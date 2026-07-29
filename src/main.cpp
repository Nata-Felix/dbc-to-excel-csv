#include <algorithm>
#include <array>
#include <cerrno>
#include <cctype>
#include <chrono>
#include <cstdint>
#include <cstdlib>
#include <filesystem>
#include <fstream>
#include <functional>
#include <iomanip>
#include <iostream>
#include <limits>
#include <memory>
#include <sstream>
#include <stdexcept>
#include <string>
#include <string_view>
#include <thread>
#include <vector>

#include "duckdb.h"

#ifdef _WIN32
#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <shellapi.h>
#include <shlobj.h>
#endif

extern "C" {
#include "blast.h"
}

namespace fs = std::filesystem;

namespace {

constexpr const char* kVersion = "1.1.0";
constexpr std::uint32_t kExcelMaxRows = 1048576;
constexpr std::uint32_t kExcelMaxColumns = 16384;

struct Options {
    fs::path input;
    fs::path output;
    std::vector<std::string> formats{"csv"};
    char delimiter = ';';
    std::string encoding = "cp1252";
    bool bom = true;
    bool include_deleted = false;
    bool recursive = false;
};

struct Field {
    std::string name;
    char type = 'C';
    std::uint8_t length = 0;
    std::uint8_t decimals = 0;
};

struct DbfSchema {
    std::uint32_t record_count = 0;
    std::uint16_t header_length = 0;
    std::uint16_t record_length = 0;
    std::uint8_t language_driver = 0;
    std::vector<Field> fields;
};

struct Cell {
    std::string text;
    char type = 'C';
    bool blank = false;
};

using RowCallback = std::function<void(std::uint32_t, const std::vector<Cell>&)>;

std::uint16_t le16(const unsigned char* p) {
    return static_cast<std::uint16_t>(p[0]) |
           (static_cast<std::uint16_t>(p[1]) << 8);
}

std::uint32_t le32(const unsigned char* p) {
    return static_cast<std::uint32_t>(p[0]) |
           (static_cast<std::uint32_t>(p[1]) << 8) |
           (static_cast<std::uint32_t>(p[2]) << 16) |
           (static_cast<std::uint32_t>(p[3]) << 24);
}

void write_le16(std::ostream& out, std::uint16_t value) {
    const unsigned char b[2] = {
        static_cast<unsigned char>(value & 0xff),
        static_cast<unsigned char>((value >> 8) & 0xff)
    };
    out.write(reinterpret_cast<const char*>(b), 2);
}

void write_le32(std::ostream& out, std::uint32_t value) {
    const unsigned char b[4] = {
        static_cast<unsigned char>(value & 0xff),
        static_cast<unsigned char>((value >> 8) & 0xff),
        static_cast<unsigned char>((value >> 16) & 0xff),
        static_cast<unsigned char>((value >> 24) & 0xff)
    };
    out.write(reinterpret_cast<const char*>(b), 4);
}

std::string lower_ascii(std::string value) {
    std::transform(value.begin(), value.end(), value.begin(), [](unsigned char c) {
        return static_cast<char>(std::tolower(c));
    });
    return value;
}

std::string trim_right(std::string value) {
    while (!value.empty() && (value.back() == ' ' || value.back() == '\0')) value.pop_back();
    return value;
}

std::string trim(std::string value) {
    value = trim_right(std::move(value));
    std::size_t pos = 0;
    while (pos < value.size() && value[pos] == ' ') ++pos;
    return value.substr(pos);
}

void append_utf8(std::string& out, std::uint32_t cp) {
    if (cp <= 0x7f) {
        out.push_back(static_cast<char>(cp));
    } else if (cp <= 0x7ff) {
        out.push_back(static_cast<char>(0xc0 | (cp >> 6)));
        out.push_back(static_cast<char>(0x80 | (cp & 0x3f)));
    } else if (cp <= 0xffff) {
        out.push_back(static_cast<char>(0xe0 | (cp >> 12)));
        out.push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3f)));
        out.push_back(static_cast<char>(0x80 | (cp & 0x3f)));
    } else {
        out.push_back(static_cast<char>(0xf0 | (cp >> 18)));
        out.push_back(static_cast<char>(0x80 | ((cp >> 12) & 0x3f)));
        out.push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3f)));
        out.push_back(static_cast<char>(0x80 | (cp & 0x3f)));
    }
}

std::string single_byte_to_utf8(std::string_view input, bool cp1252) {
    static constexpr std::array<std::uint16_t, 32> cp1252_map = {
        0x20ac, 0xfffd, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021,
        0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0xfffd, 0x017d, 0xfffd,
        0xfffd, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
        0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0xfffd, 0x017e, 0x0178
    };
    std::string out;
    out.reserve(input.size() * 2);
    for (unsigned char c : input) {
        std::uint32_t cp = c;
        if (cp1252 && c >= 0x80 && c <= 0x9f) cp = cp1252_map[c - 0x80];
        if (cp == 0) continue;
        append_utf8(out, cp);
    }
    return out;
}

std::string decode_text(std::string_view input, const std::string& encoding) {
    if (encoding == "utf8") return std::string(input);
    return single_byte_to_utf8(input, encoding == "cp1252");
}

std::string normalize_date(std::string value) {
    if (value.size() == 8 && std::all_of(value.begin(), value.end(), [](unsigned char c) { return std::isdigit(c); })) {
        return value.substr(0, 4) + "-" + value.substr(4, 2) + "-" + value.substr(6, 2);
    }
    return value;
}

std::string format_cell(std::string raw, char type, const std::string& encoding) {
    switch (type) {
        case 'C':
        case 'M':
            return decode_text(trim_right(std::move(raw)), encoding);
        case 'D':
            return normalize_date(trim(std::move(raw)));
        default:
            return decode_text(trim(std::move(raw)), encoding);
    }
}

struct BlastInput {
    std::ifstream* stream = nullptr;
    std::array<unsigned char, 4096> buffer{};
};

struct BlastOutput {
    std::ofstream* stream = nullptr;
};

unsigned blast_input(void* how, unsigned char** buffer) {
    auto* input = static_cast<BlastInput*>(how);
    input->stream->read(reinterpret_cast<char*>(input->buffer.data()),
                        static_cast<std::streamsize>(input->buffer.size()));
    *buffer = input->buffer.data();
    return static_cast<unsigned>(input->stream->gcount());
}

int blast_output(void* how, unsigned char* buffer, unsigned length) {
    auto* output = static_cast<BlastOutput*>(how);
    output->stream->write(reinterpret_cast<const char*>(buffer), length);
    return output->stream->good() ? 0 : 1;
}

void decompress_dbc(const fs::path& input_path, const fs::path& output_path) {
    std::ifstream input(input_path, std::ios::binary);
    if (!input) throw std::runtime_error("nao foi possivel abrir o DBC de entrada");

    std::array<unsigned char, 32> fixed{};
    input.read(reinterpret_cast<char*>(fixed.data()), fixed.size());
    if (input.gcount() != static_cast<std::streamsize>(fixed.size())) {
        throw std::runtime_error("arquivo pequeno demais para ser um DBC/DBF valido");
    }
    const std::uint16_t header_length = le16(fixed.data() + 8);
    const std::uint16_t record_length = le16(fixed.data() + 10);
    if (header_length < 33 || record_length < 1 || header_length > 65535) {
        throw std::runtime_error("cabecalho DBC invalido");
    }

    input.seekg(0, std::ios::beg);
    std::vector<unsigned char> header(header_length);
    input.read(reinterpret_cast<char*>(header.data()), header.size());
    if (input.gcount() != static_cast<std::streamsize>(header.size())) {
        throw std::runtime_error("cabecalho DBC truncado");
    }
    header.back() = 0x0d;

    std::ofstream output(output_path, std::ios::binary | std::ios::trunc);
    if (!output) throw std::runtime_error("nao foi possivel criar o DBF de trabalho");
    output.write(reinterpret_cast<const char*>(header.data()), header.size());

    input.seekg(static_cast<std::streamoff>(header_length) + 4, std::ios::beg);
    if (!input) throw std::runtime_error("DBC truncado antes dos dados comprimidos");

    BlastInput blast_in{&input};
    BlastOutput blast_out{&output};
    unsigned left = 0;
    unsigned char* unused = nullptr;
    const int result = blast(blast_input, &blast_in, blast_output, &blast_out, &left, &unused);
    output.flush();
    if (result != 0) {
        throw std::runtime_error("falha blast ao descomprimir DBC (codigo " + std::to_string(result) + ")");
    }
    if (!output) throw std::runtime_error("erro ao gravar o DBF descomprimido");
}

DbfSchema read_schema(std::ifstream& input) {
    input.clear();
    input.seekg(0, std::ios::beg);
    std::array<unsigned char, 32> header{};
    input.read(reinterpret_cast<char*>(header.data()), header.size());
    if (input.gcount() != static_cast<std::streamsize>(header.size())) {
        throw std::runtime_error("DBF truncado");
    }

    DbfSchema schema;
    schema.record_count = le32(header.data() + 4);
    schema.header_length = le16(header.data() + 8);
    schema.record_length = le16(header.data() + 10);
    schema.language_driver = header[29];
    if (schema.header_length < 33 || schema.record_length < 1) {
        throw std::runtime_error("estrutura DBF invalida");
    }

    const std::size_t descriptor_bytes = schema.header_length - 33;
    if (descriptor_bytes % 32 != 0) {
        throw std::runtime_error("tamanho do cabecalho DBF inconsistente");
    }
    const std::size_t field_count = descriptor_bytes / 32;
    for (std::size_t i = 0; i < field_count; ++i) {
        std::array<unsigned char, 32> descriptor{};
        input.read(reinterpret_cast<char*>(descriptor.data()), descriptor.size());
        if (!input) throw std::runtime_error("descritor de campo DBF truncado");
        std::size_t name_length = 0;
        while (name_length < 11 && descriptor[name_length] != 0) ++name_length;
        Field field;
        field.name.assign(reinterpret_cast<const char*>(descriptor.data()), name_length);
        field.type = static_cast<char>(descriptor[11]);
        field.length = descriptor[16];
        field.decimals = descriptor[17];
        schema.fields.push_back(std::move(field));
    }

    char terminator = 0;
    input.get(terminator);
    if (!input || static_cast<unsigned char>(terminator) != 0x0d) {
        throw std::runtime_error("terminador do cabecalho DBF ausente");
    }
    std::uint32_t calculated = 1;
    for (const auto& field : schema.fields) calculated += field.length;
    if (calculated > schema.record_length) {
        throw std::runtime_error("campos DBF excedem o tamanho do registro");
    }
    return schema;
}

std::uint32_t read_rows(const fs::path& dbf_path, const DbfSchema& schema,
                        const Options& options, const RowCallback& callback) {
    std::ifstream input(dbf_path, std::ios::binary);
    if (!input) throw std::runtime_error("nao foi possivel reler o DBF");
    input.seekg(schema.header_length, std::ios::beg);
    std::vector<char> record(schema.record_length);
    std::uint32_t emitted = 0;
    for (std::uint32_t record_index = 0; record_index < schema.record_count; ++record_index) {
        input.read(record.data(), record.size());
        if (input.gcount() != static_cast<std::streamsize>(record.size())) {
            throw std::runtime_error("DBF truncado no registro " + std::to_string(record_index + 1));
        }
        const bool deleted = record[0] == '*';
        if (deleted && !options.include_deleted) continue;
        std::vector<Cell> cells;
        cells.reserve(schema.fields.size());
        std::size_t offset = 1;
        for (const auto& field : schema.fields) {
            std::string raw(record.data() + offset, field.length);
            std::string text = format_cell(std::move(raw), field.type, options.encoding);
            cells.push_back(Cell{std::move(text), field.type, false});
            cells.back().blank = cells.back().text.empty();
            offset += field.length;
        }
        callback(emitted, cells);
        ++emitted;
    }
    return emitted;
}

std::string csv_escape(const std::string& value, char delimiter) {
    const bool quote = value.find(delimiter) != std::string::npos ||
                       value.find('"') != std::string::npos ||
                       value.find('\r') != std::string::npos ||
                       value.find('\n') != std::string::npos;
    if (!quote) return value;
    std::string out = "\"";
    for (char c : value) {
        if (c == '"') out += "\"\"";
        else out.push_back(c);
    }
    out.push_back('"');
    return out;
}

void write_csv(const fs::path& dbf_path, const fs::path& output_path,
               const DbfSchema& schema, const Options& options) {
    std::ofstream out(output_path, std::ios::binary | std::ios::trunc);
    if (!out) throw std::runtime_error("nao foi possivel criar o CSV");
    if (options.bom) out.write("\xef\xbb\xbf", 3);
    for (std::size_t i = 0; i < schema.fields.size(); ++i) {
        if (i) out.put(options.delimiter);
        out << csv_escape(schema.fields[i].name, options.delimiter);
    }
    out << "\r\n";
    read_rows(dbf_path, schema, options, [&](std::uint32_t, const std::vector<Cell>& cells) {
        for (std::size_t i = 0; i < cells.size(); ++i) {
            if (i) out.put(options.delimiter);
            out << csv_escape(cells[i].text, options.delimiter);
        }
        out << "\r\n";
        if (!out) throw std::runtime_error("erro ao gravar CSV");
    });
}

std::string json_escape(std::string_view value) {
    std::string out;
    out.reserve(value.size() + 8);
    for (unsigned char c : value) {
        switch (c) {
            case '"': out += "\\\""; break;
            case '\\': out += "\\\\"; break;
            case '\b': out += "\\b"; break;
            case '\f': out += "\\f"; break;
            case '\n': out += "\\n"; break;
            case '\r': out += "\\r"; break;
            case '\t': out += "\\t"; break;
            default:
                if (c < 0x20) {
                    const char hex[] = "0123456789abcdef";
                    out += "\\u00";
                    out.push_back(hex[c >> 4]);
                    out.push_back(hex[c & 0x0f]);
                } else {
                    out.push_back(static_cast<char>(c));
                }
        }
    }
    return out;
}

bool valid_json_number(const std::string& value) {
    if (value.empty()) return false;
    std::size_t i = 0;
    if (value[i] == '-') ++i;
    if (i >= value.size()) return false;
    if (value[i] == '0') {
        ++i;
    } else if (value[i] >= '1' && value[i] <= '9') {
        while (i < value.size() && std::isdigit(static_cast<unsigned char>(value[i]))) ++i;
    } else return false;
    if (i < value.size() && value[i] == '.') {
        ++i;
        const std::size_t start = i;
        while (i < value.size() && std::isdigit(static_cast<unsigned char>(value[i]))) ++i;
        if (i == start) return false;
    }
    if (i < value.size() && (value[i] == 'e' || value[i] == 'E')) {
        ++i;
        if (i < value.size() && (value[i] == '+' || value[i] == '-')) ++i;
        const std::size_t start = i;
        while (i < value.size() && std::isdigit(static_cast<unsigned char>(value[i]))) ++i;
        if (i == start) return false;
    }
    return i == value.size();
}

void write_jsonl(const fs::path& dbf_path, const fs::path& output_path,
                 const DbfSchema& schema, const Options& options) {
    std::ofstream out(output_path, std::ios::binary | std::ios::trunc);
    if (!out) throw std::runtime_error("nao foi possivel criar o JSONL");
    read_rows(dbf_path, schema, options, [&](std::uint32_t, const std::vector<Cell>& cells) {
        out.put('{');
        for (std::size_t i = 0; i < cells.size(); ++i) {
            if (i) out.put(',');
            out << '"' << json_escape(schema.fields[i].name) << "\":";
            const Cell& cell = cells[i];
            if (cell.blank) {
                out << "null";
            } else if ((cell.type == 'N' || cell.type == 'F' || cell.type == 'I' || cell.type == 'Y') &&
                       valid_json_number(cell.text)) {
                out << cell.text;
            } else if (cell.type == 'L') {
                const char v = static_cast<char>(std::toupper(static_cast<unsigned char>(cell.text[0])));
                if (v == 'T' || v == 'Y') out << "true";
                else if (v == 'F' || v == 'N') out << "false";
                else out << "null";
            } else {
                out << '"' << json_escape(cell.text) << '"';
            }
        }
        out << "}\n";
        if (!out) throw std::runtime_error("erro ao gravar JSONL");
    });
}

std::uint32_t crc32_update(std::uint32_t crc, const unsigned char* data, std::size_t length) {
    static const std::array<std::uint32_t, 256> table = [] {
        std::array<std::uint32_t, 256> values{};
        for (std::uint32_t i = 0; i < 256; ++i) {
            std::uint32_t c = i;
            for (int bit = 0; bit < 8; ++bit) c = (c & 1) ? 0xedb88320U ^ (c >> 1) : c >> 1;
            values[i] = c;
        }
        return values;
    }();
    crc ^= 0xffffffffU;
    for (std::size_t i = 0; i < length; ++i) crc = table[(crc ^ data[i]) & 0xff] ^ (crc >> 8);
    return crc ^ 0xffffffffU;
}

class ZipWriter {
public:
    explicit ZipWriter(const fs::path& path) : out_(path, std::ios::binary | std::ios::trunc) {
        if (!out_) throw std::runtime_error("nao foi possivel criar o XLSX");
    }

    class EntryStream {
    public:
        explicit EntryStream(ZipWriter& owner) : owner_(owner) {}
        void write(std::string_view data) {
            if (data.empty()) return;
            owner_.out_.write(data.data(), static_cast<std::streamsize>(data.size()));
            if (!owner_.out_) throw std::runtime_error("erro ao gravar XLSX");
            crc_ = crc32_update(crc_, reinterpret_cast<const unsigned char*>(data.data()), data.size());
            size_ += data.size();
            if (size_ > std::numeric_limits<std::uint32_t>::max()) {
                throw std::runtime_error("uma parte do XLSX excedeu o limite ZIP32 de 4 GB");
            }
        }
        std::uint32_t crc() const { return crc_; }
        std::uint32_t size() const { return static_cast<std::uint32_t>(size_); }
    private:
        ZipWriter& owner_;
        std::uint32_t crc_ = 0;
        std::uint64_t size_ = 0;
    };

    void add(const std::string& name, const std::function<void(EntryStream&)>& producer) {
        const auto offset64 = static_cast<std::uint64_t>(out_.tellp());
        if (offset64 > std::numeric_limits<std::uint32_t>::max()) {
            throw std::runtime_error("XLSX excedeu o limite ZIP32 de 4 GB");
        }
        Entry entry;
        entry.name = name;
        entry.offset = static_cast<std::uint32_t>(offset64);

        write_le32(out_, 0x04034b50);
        write_le16(out_, 20);
        write_le16(out_, 0x0808);
        write_le16(out_, 0);
        write_le16(out_, 0);
        write_le16(out_, 0);
        write_le32(out_, 0);
        write_le32(out_, 0);
        write_le32(out_, 0);
        write_le16(out_, static_cast<std::uint16_t>(name.size()));
        write_le16(out_, 0);
        out_.write(name.data(), static_cast<std::streamsize>(name.size()));

        EntryStream stream(*this);
        producer(stream);
        entry.crc = stream.crc();
        entry.size = stream.size();
        write_le32(out_, 0x08074b50);
        write_le32(out_, entry.crc);
        write_le32(out_, entry.size);
        write_le32(out_, entry.size);
        entries_.push_back(std::move(entry));
    }

    void finish() {
        if (finished_) return;
        const auto central_offset64 = static_cast<std::uint64_t>(out_.tellp());
        if (central_offset64 > std::numeric_limits<std::uint32_t>::max()) {
            throw std::runtime_error("XLSX excedeu o limite ZIP32 de 4 GB");
        }
        const auto central_offset = static_cast<std::uint32_t>(central_offset64);
        if (entries_.size() > std::numeric_limits<std::uint16_t>::max()) {
            throw std::runtime_error("XLSX contem partes demais");
        }
        for (const auto& entry : entries_) {
            write_le32(out_, 0x02014b50);
            write_le16(out_, 20);
            write_le16(out_, 20);
            write_le16(out_, 0x0808);
            write_le16(out_, 0);
            write_le16(out_, 0);
            write_le16(out_, 0);
            write_le32(out_, entry.crc);
            write_le32(out_, entry.size);
            write_le32(out_, entry.size);
            write_le16(out_, static_cast<std::uint16_t>(entry.name.size()));
            write_le16(out_, 0);
            write_le16(out_, 0);
            write_le16(out_, 0);
            write_le16(out_, 0);
            write_le32(out_, 0);
            write_le32(out_, entry.offset);
            out_.write(entry.name.data(), static_cast<std::streamsize>(entry.name.size()));
        }
        const auto central_end64 = static_cast<std::uint64_t>(out_.tellp());
        const auto central_size64 = central_end64 - central_offset64;
        if (central_size64 > std::numeric_limits<std::uint32_t>::max()) {
            throw std::runtime_error("diretorio central XLSX grande demais");
        }
        write_le32(out_, 0x06054b50);
        write_le16(out_, 0);
        write_le16(out_, 0);
        write_le16(out_, static_cast<std::uint16_t>(entries_.size()));
        write_le16(out_, static_cast<std::uint16_t>(entries_.size()));
        write_le32(out_, static_cast<std::uint32_t>(central_size64));
        write_le32(out_, central_offset);
        write_le16(out_, 0);
        out_.flush();
        if (!out_) throw std::runtime_error("erro ao finalizar XLSX");
        finished_ = true;
    }

    ~ZipWriter() {
        if (!finished_) {
            try { finish(); } catch (...) {}
        }
    }

private:
    struct Entry {
        std::string name;
        std::uint32_t crc = 0;
        std::uint32_t size = 0;
        std::uint32_t offset = 0;
    };
    std::ofstream out_;
    std::vector<Entry> entries_;
    bool finished_ = false;
};

std::string xml_escape(std::string_view value) {
    std::string out;
    out.reserve(value.size() + 16);
    for (unsigned char c : value) {
        if (c < 0x20 && c != '\t' && c != '\n' && c != '\r') continue;
        switch (c) {
            case '&': out += "&amp;"; break;
            case '<': out += "&lt;"; break;
            case '>': out += "&gt;"; break;
            case '"': out += "&quot;"; break;
            case '\'': out += "&apos;"; break;
            default: out.push_back(static_cast<char>(c));
        }
    }
    return out;
}

std::string excel_column(std::size_t index) {
    std::string result;
    ++index;
    while (index) {
        const std::size_t rem = (index - 1) % 26;
        result.push_back(static_cast<char>('A' + rem));
        index = (index - 1) / 26;
    }
    std::reverse(result.begin(), result.end());
    return result;
}

bool excel_safe_number(const Cell& cell) {
    if (!(cell.type == 'N' || cell.type == 'F' || cell.type == 'I' || cell.type == 'Y')) return false;
    if (!valid_json_number(cell.text)) return false;
    std::size_t digits = 0;
    for (unsigned char c : cell.text) if (std::isdigit(c)) ++digits;
    return digits <= 15;
}

std::int64_t days_from_civil(int year, unsigned month, unsigned day) {
    year -= month <= 2;
    const int era = (year >= 0 ? year : year - 399) / 400;
    const unsigned yoe = static_cast<unsigned>(year - era * 400);
    const unsigned doy = (153 * (month + (month > 2 ? -3 : 9)) + 2) / 5 + day - 1;
    const unsigned doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    return static_cast<std::int64_t>(era) * 146097 + static_cast<std::int64_t>(doe) - 719468;
}

bool excel_date_serial(const Cell& cell, std::int64_t& serial) {
    if (cell.type != 'D' || cell.text.size() != 10 || cell.text[4] != '-' || cell.text[7] != '-') return false;
    try {
        const int year = std::stoi(cell.text.substr(0, 4));
        const unsigned month = static_cast<unsigned>(std::stoi(cell.text.substr(5, 2)));
        const unsigned day = static_cast<unsigned>(std::stoi(cell.text.substr(8, 2)));
        if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return false;
        serial = days_from_civil(year, month, day) + 25569;
        return serial > 0;
    } catch (...) {
        return false;
    }
}

std::string sql_identifier(const std::string& value) {
    std::string out = "\"";
    for (char c : value) {
        if (c == '"') out += "\"\"";
        else out.push_back(c);
    }
    out.push_back('"');
    return out;
}

std::string sql_literal(const std::string& value) {
    std::string out = "'";
    for (char c : value) {
        if (c == '\'') out += "\'\'";
        else out.push_back(c);
    }
    out.push_back('\'');
    return out;
}

void duckdb_exec(duckdb_connection connection, const std::string& sql) {
    duckdb_result result{};
    if (duckdb_query(connection, sql.c_str(), &result) == DuckDBError) {
        const char* message = duckdb_result_error(&result);
        const std::string detail = message ? message : "erro desconhecido";
        duckdb_destroy_result(&result);
        throw std::runtime_error("DuckDB/Parquet: " + detail);
    }
    duckdb_destroy_result(&result);
}

std::string parquet_sql_type(const Field& field) {
    if (field.type == 'D') return "DATE";
    if (field.type == 'L') return "BOOLEAN";
    if ((field.type == 'N' || field.type == 'F') && field.decimals == 0 && field.length <= 18) return "BIGINT";
    if (field.type == 'N' || field.type == 'F') return "DOUBLE";
    return "VARCHAR";
}

void write_parquet(const fs::path& dbf_path, const fs::path& output_path,
                   const DbfSchema& schema, const Options& options) {
    const fs::path database_path = output_path.string() + ".duckdb.tmp";
    struct DatabaseFilesGuard {
        fs::path path;
        ~DatabaseFilesGuard() {
            std::error_code ec;
            fs::remove(path, ec);
            fs::remove(path.string() + ".wal", ec);
        }
    } files_guard{database_path};

    duckdb_database database = nullptr;
    duckdb_connection connection = nullptr;
    duckdb_appender appender = nullptr;
    struct DuckGuard {
        duckdb_database* database;
        duckdb_connection* connection;
        duckdb_appender* appender;
        ~DuckGuard() {
            if (*appender) duckdb_appender_destroy(appender);
            if (*connection) duckdb_disconnect(connection);
            if (*database) duckdb_close(database);
        }
    } duck_guard{&database, &connection, &appender};

    const std::string database_utf8 = database_path.u8string();
    if (duckdb_open(database_utf8.c_str(), &database) == DuckDBError) {
        throw std::runtime_error("DuckDB/Parquet: nao foi possivel criar o banco temporario");
    }
    if (duckdb_connect(database, &connection) == DuckDBError) {
        throw std::runtime_error("DuckDB/Parquet: nao foi possivel iniciar a conexao");
    }

    std::string create = "CREATE TABLE dados (";
    for (std::size_t i = 0; i < schema.fields.size(); ++i) {
        if (i) create += ',';
        create += sql_identifier(schema.fields[i].name) + " " + parquet_sql_type(schema.fields[i]);
    }
    create += ")";
    duckdb_exec(connection, create);

    if (duckdb_appender_create(connection, nullptr, "dados", &appender) == DuckDBError) {
        throw std::runtime_error("DuckDB/Parquet: falha ao iniciar a carga de registros");
    }
    read_rows(dbf_path, schema, options, [&](std::uint32_t row_index, const std::vector<Cell>& cells) {
        if (duckdb_appender_begin_row(appender) == DuckDBError) {
            throw std::runtime_error("DuckDB/Parquet: falha no registro " + std::to_string(row_index + 1));
        }
        for (std::size_t i = 0; i < cells.size(); ++i) {
            const Cell& cell = cells[i];
            const Field& field = schema.fields[i];
            duckdb_state state = DuckDBSuccess;
            if (cell.blank) {
                state = duckdb_append_null(appender);
            } else if ((field.type == 'N' || field.type == 'F') && field.decimals == 0 && field.length <= 18) {
                errno = 0;
                char* end = nullptr;
                const long long value = std::strtoll(cell.text.c_str(), &end, 10);
                if (errno == 0 && end && *end == '\0') state = duckdb_append_int64(appender, static_cast<std::int64_t>(value));
                else state = duckdb_append_null(appender);
            } else if (field.type == 'N' || field.type == 'F') {
                errno = 0;
                char* end = nullptr;
                const double value = std::strtod(cell.text.c_str(), &end);
                if (errno == 0 && end && *end == '\0') state = duckdb_append_double(appender, value);
                else state = duckdb_append_null(appender);
            } else if (field.type == 'D') {
                std::int64_t serial = 0;
                if (excel_date_serial(cell, serial)) {
                    duckdb_date date{};
                    date.days = static_cast<std::int32_t>(serial - 25569);
                    state = duckdb_append_date(appender, date);
                } else {
                    state = duckdb_append_null(appender);
                }
            } else if (field.type == 'L') {
                const char value = static_cast<char>(std::toupper(static_cast<unsigned char>(cell.text[0])));
                if (value == 'T' || value == 'Y') state = duckdb_append_bool(appender, true);
                else if (value == 'F' || value == 'N') state = duckdb_append_bool(appender, false);
                else state = duckdb_append_null(appender);
            } else {
                state = duckdb_append_varchar_length(appender, cell.text.data(), cell.text.size());
            }
            if (state == DuckDBError) {
                const char* detail = duckdb_appender_error(appender);
                throw std::runtime_error("DuckDB/Parquet no campo " + field.name + ": " +
                                         (detail ? detail : "falha ao anexar valor"));
            }
        }
        if (duckdb_appender_end_row(appender) == DuckDBError) {
            const char* detail = duckdb_appender_error(appender);
            throw std::runtime_error("DuckDB/Parquet ao finalizar registro: " +
                                     std::string(detail ? detail : "erro desconhecido"));
        }
    });
    if (duckdb_appender_close(appender) == DuckDBError) {
        const char* detail = duckdb_appender_error(appender);
        throw std::runtime_error("DuckDB/Parquet ao finalizar carga: " +
                                 std::string(detail ? detail : "erro desconhecido"));
    }
    duckdb_appender_destroy(&appender);

    const std::string output_utf8 = output_path.u8string();
    duckdb_exec(connection, "COPY dados TO " + sql_literal(output_utf8) +
                            " (FORMAT PARQUET, COMPRESSION ZSTD, COMPRESSION_LEVEL 3)");
}

void write_inline_string(ZipWriter::EntryStream& stream, const std::string& reference,
                         const std::string& value, bool header = false) {
    std::string cell = "<c r=\"" + reference + "\" t=\"inlineStr\"";
    if (header) cell += " s=\"1\"";
    cell += "><is><t xml:space=\"preserve\">" + xml_escape(value) + "</t></is></c>";
    stream.write(cell);
}

void write_xlsx(const fs::path& dbf_path, const fs::path& output_path,
                const DbfSchema& schema, const Options& options) {
    if (schema.fields.size() > kExcelMaxColumns) {
        throw std::runtime_error("o DBF tem mais de 16.384 colunas, limite do Excel");
    }
    ZipWriter zip(output_path);
    zip.add("[Content_Types].xml", [](auto& s) {
        s.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
                "</Types>");
    });
    zip.add("_rels/.rels", [](auto& s) {
        s.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                "</Relationships>");
    });
    zip.add("xl/workbook.xml", [](auto& s) {
        s.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                "<sheets><sheet name=\"Dados\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
    });
    zip.add("xl/_rels/workbook.xml.rels", [](auto& s) {
        s.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
                "</Relationships>");
    });
    zip.add("xl/styles.xml", [](auto& s) {
        s.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                "<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd\"/></numFmts>"
                "<fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font>"
                "<font><b/><color rgb=\"FFFFFFFF\"/><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
                "<fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>"
                "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF1F4E78\"/><bgColor indexed=\"64\"/></patternFill></fill></fills>"
                "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
                "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
                "<cellXfs count=\"3\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
                "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/>"
                "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/></cellXfs>"
                "</styleSheet>");
    });
    zip.add("xl/worksheets/sheet1.xml", [&](auto& s) {
        s.write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                "<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>"
                "<cols>");
        for (std::size_t i = 0; i < schema.fields.size(); ++i) {
            const auto& field = schema.fields[i];
            const std::size_t semantic = field.type == 'D' ? 12 : static_cast<std::size_t>(field.length) + 1;
            const std::size_t width = std::min<std::size_t>(30, std::max<std::size_t>(field.name.size() + 2, semantic));
            s.write("<col min=\"" + std::to_string(i + 1) + "\" max=\"" + std::to_string(i + 1) +
                    "\" width=\"" + std::to_string(width) + "\" customWidth=\"1\"/>");
        }
        s.write("</cols><sheetData><row r=\"1\">");
        for (std::size_t i = 0; i < schema.fields.size(); ++i) {
            write_inline_string(s, excel_column(i) + "1", schema.fields[i].name, true);
        }
        s.write("</row>");
        const std::uint32_t written_rows = read_rows(dbf_path, schema, options, [&](std::uint32_t row_index, const std::vector<Cell>& cells) {
            const std::uint32_t excel_row = row_index + 2;
            if (excel_row > kExcelMaxRows) {
                throw std::runtime_error("mais de 1.048.575 registros: excede o limite de linhas do Excel; use CSV ou JSONL");
            }
            s.write("<row r=\"" + std::to_string(excel_row) + "\">");
            for (std::size_t i = 0; i < cells.size(); ++i) {
                const Cell& cell = cells[i];
                if (cell.blank) continue;
                const std::string ref = excel_column(i) + std::to_string(excel_row);
                std::int64_t date_serial = 0;
                if (excel_date_serial(cell, date_serial)) {
                    s.write("<c r=\"" + ref + "\" s=\"2\"><v>" + std::to_string(date_serial) + "</v></c>");
                } else if (excel_safe_number(cell)) {
                    s.write("<c r=\"" + ref + "\"><v>" + cell.text + "</v></c>");
                } else if (cell.type == 'L') {
                    const char v = static_cast<char>(std::toupper(static_cast<unsigned char>(cell.text[0])));
                    if (v == 'T' || v == 'Y') s.write("<c r=\"" + ref + "\" t=\"b\"><v>1</v></c>");
                    else if (v == 'F' || v == 'N') s.write("<c r=\"" + ref + "\" t=\"b\"><v>0</v></c>");
                    else write_inline_string(s, ref, cell.text);
                } else {
                    write_inline_string(s, ref, cell.text);
                }
            }
            s.write("</row>");
        });
        s.write("</sheetData><autoFilter ref=\"A1:" + excel_column(schema.fields.size() - 1) +
                std::to_string(written_rows + 1) + "\"/></worksheet>");
    });
    zip.finish();
}

fs::path unique_temp_dbf(const fs::path& directory, const std::string& stem) {
    const auto ticks = std::chrono::high_resolution_clock::now().time_since_epoch().count();
    return directory / ("." + stem + ".dbcconv-" + std::to_string(ticks) + ".tmp.dbf");
}

void replace_file(const fs::path& source, const fs::path& destination) {
    std::error_code ec;
    fs::remove(destination, ec);
    ec.clear();
    fs::rename(source, destination, ec);
    if (ec) throw std::runtime_error("nao foi possivel finalizar " + destination.filename().string() + ": " + ec.message());
}

std::vector<std::string> expand_formats(const std::vector<std::string>& requested) {
    std::vector<std::string> result;
    for (const auto& item : requested) {
        if (item == "all") {
            for (const char* format : {"dbf", "csv", "xlsx", "parquet", "jsonl"}) {
                if (std::find(result.begin(), result.end(), format) == result.end()) result.emplace_back(format);
            }
        } else if (std::find(result.begin(), result.end(), item) == result.end()) {
            result.push_back(item);
        }
    }
    return result;
}

std::uint32_t convert_one(const fs::path& input, const fs::path& output_dir, const Options& options) {
    fs::create_directories(output_dir);
    const std::string stem = input.stem().string();
    const auto formats = expand_formats(options.formats);
    const fs::path temp_dbf = unique_temp_dbf(output_dir, stem);
    struct TempGuard {
        fs::path path;
        ~TempGuard() { std::error_code ec; fs::remove(path, ec); }
    } guard{temp_dbf};

    std::cout << "[DBC] " << input.u8string() << "\n";
    decompress_dbc(input, temp_dbf);
    std::ifstream dbf_input(temp_dbf, std::ios::binary);
    if (!dbf_input) throw std::runtime_error("nao foi possivel abrir o DBF de trabalho");
    const DbfSchema schema = read_schema(dbf_input);
    dbf_input.close();

    for (const auto& format : formats) {
        const fs::path final_path = output_dir / (stem + "." + format);
        const fs::path part_path = final_path.string() + ".part";
        std::error_code cleanup;
        fs::remove(part_path, cleanup);
        try {
            if (format == "dbf") {
                fs::copy_file(temp_dbf, part_path, fs::copy_options::overwrite_existing);
            } else if (format == "csv") {
                write_csv(temp_dbf, part_path, schema, options);
            } else if (format == "xlsx") {
                write_xlsx(temp_dbf, part_path, schema, options);
            } else if (format == "parquet") {
                write_parquet(temp_dbf, part_path, schema, options);
            } else if (format == "jsonl") {
                write_jsonl(temp_dbf, part_path, schema, options);
            } else {
                throw std::runtime_error("formato desconhecido: " + format);
            }
            replace_file(part_path, final_path);
            std::cout << "  OK  " << final_path.u8string() << "\n";
        } catch (...) {
            fs::remove(part_path, cleanup);
            throw;
        }
    }
    std::cout << "  " << schema.record_count << " registros declarados, "
              << schema.fields.size() << " colunas\n";
    return schema.record_count;
}

void print_help(const char* program) {
    std::cout
        << "DBC Converter " << kVersion << " - DATASUS DBC para DBF/CSV/XLSX/Parquet/JSONL\n\n"
        << "Uso:\n  " << program << " <arquivo.dbc|pasta> [opcoes]\n\n"
        << "Opcoes:\n"
        << "  -f, --format <tipo>       csv, xlsx, parquet, dbf, jsonl ou all (padrao: csv)\n"
        << "  -o, --output <caminho>    arquivo de saida (entrada unica e um formato)\n"
        << "                             ou pasta de saida\n"
        << "  --delimiter <valor>       semicolon, comma, tab ou um caractere\n"
        << "  --encoding <tipo>         cp1252, latin1 ou utf8 (padrao: cp1252)\n"
        << "  --no-bom                  CSV UTF-8 sem BOM\n"
        << "  --include-deleted         inclui registros DBF marcados como excluidos\n"
        << "  -r, --recursive           procura DBC em subpastas\n"
        << "  -h, --help                mostra esta ajuda\n"
        << "  -v, --version             mostra a versao\n\n"
        << "Exemplos:\n"
        << "  " << program << " MENTBR26.DBC -f xlsx\n"
        << "  " << program << " C:\\dados -f all -o C:\\convertidos\n";
}

Options parse_options(const std::vector<std::string>& args) {
    Options options;
    bool format_seen = false;
    for (std::size_t i = 1; i < args.size(); ++i) {
        const std::string arg = args[i];
        auto require_value = [&](const char* name) -> std::string {
            if (++i >= args.size()) throw std::runtime_error(std::string("faltou valor para ") + name);
            return args[i];
        };
        if (arg == "-f" || arg == "--format") {
            std::string value = lower_ascii(require_value(arg.c_str()));
            if (value != "csv" && value != "xlsx" && value != "parquet" && value != "dbf" && value != "jsonl" && value != "all") {
                throw std::runtime_error("formato invalido: " + value);
            }
            if (!format_seen) { options.formats.clear(); format_seen = true; }
            options.formats.push_back(value);
        } else if (arg == "-o" || arg == "--output") {
            options.output = fs::u8path(require_value(arg.c_str()));
        } else if (arg == "--delimiter") {
            std::string value = lower_ascii(require_value(arg.c_str()));
            if (value == "semicolon") options.delimiter = ';';
            else if (value == "comma") options.delimiter = ',';
            else if (value == "tab") options.delimiter = '\t';
            else if (value.size() == 1) options.delimiter = value[0];
            else throw std::runtime_error("delimitador invalido");
        } else if (arg == "--encoding") {
            options.encoding = lower_ascii(require_value(arg.c_str()));
            if (options.encoding == "windows-1252") options.encoding = "cp1252";
            if (options.encoding == "iso-8859-1") options.encoding = "latin1";
            if (options.encoding == "utf-8") options.encoding = "utf8";
            if (options.encoding != "cp1252" && options.encoding != "latin1" && options.encoding != "utf8") {
                throw std::runtime_error("encoding invalido: use cp1252, latin1 ou utf8");
            }
        } else if (arg == "--no-bom") {
            options.bom = false;
        } else if (arg == "--include-deleted") {
            options.include_deleted = true;
        } else if (arg == "-r" || arg == "--recursive") {
            options.recursive = true;
        } else if (!arg.empty() && arg[0] == '-') {
            throw std::runtime_error("opcao desconhecida: " + arg);
        } else if (options.input.empty()) {
            options.input = fs::u8path(arg);
        } else {
            throw std::runtime_error("somente uma entrada pode ser informada");
        }
    }
    if (options.input.empty()) throw std::runtime_error("informe um arquivo .dbc ou uma pasta");
    return options;
}

int run(const std::vector<std::string>& args) {
    if (args.size() == 1) {
        print_help(args[0].c_str());
        return 1;
    }
    if (args[1] == "-h" || args[1] == "--help") { print_help(args[0].c_str()); return 0; }
    if (args[1] == "-v" || args[1] == "--version") { std::cout << kVersion << "\n"; return 0; }
    const Options options = parse_options(args);
    if (!fs::exists(options.input)) throw std::runtime_error("entrada nao encontrada");

    std::vector<fs::path> inputs;
    if (fs::is_regular_file(options.input)) {
        if (lower_ascii(options.input.extension().string()) != ".dbc") {
            throw std::runtime_error("o arquivo de entrada precisa ter extensao .dbc");
        }
        inputs.push_back(options.input);
    } else if (fs::is_directory(options.input)) {
        if (options.recursive) {
            for (const auto& item : fs::recursive_directory_iterator(options.input)) {
                if (item.is_regular_file() && lower_ascii(item.path().extension().string()) == ".dbc") inputs.push_back(item.path());
            }
        } else {
            for (const auto& item : fs::directory_iterator(options.input)) {
                if (item.is_regular_file() && lower_ascii(item.path().extension().string()) == ".dbc") inputs.push_back(item.path());
            }
        }
        std::sort(inputs.begin(), inputs.end());
    } else {
        throw std::runtime_error("a entrada nao e arquivo nem pasta");
    }
    if (inputs.empty()) throw std::runtime_error("nenhum arquivo .dbc encontrado");

    const auto formats = expand_formats(options.formats);
    if (inputs.size() == 1 && formats.size() == 1 && !options.output.empty() && options.output.has_extension()) {
        const std::string wanted_ext = "." + formats[0];
        if (lower_ascii(options.output.extension().string()) != wanted_ext) {
            throw std::runtime_error("a extensao da saida nao corresponde ao formato " + formats[0]);
        }
        Options adjusted = options;
        const fs::path target_dir = options.output.has_parent_path() ? options.output.parent_path() : fs::current_path();
        fs::create_directories(target_dir);
        const fs::path expected = target_dir / (inputs[0].stem().string() + wanted_ext);
        convert_one(inputs[0], target_dir, adjusted);
        if (expected != options.output) {
            std::error_code ec;
            fs::remove(options.output, ec);
            fs::rename(expected, options.output, ec);
            if (ec) throw std::runtime_error("nao foi possivel renomear a saida: " + ec.message());
        }
    } else {
        fs::path output_dir;
        if (!options.output.empty()) output_dir = options.output;
        else if (fs::is_directory(options.input)) output_dir = options.input / "convertidos";
        else output_dir = options.input.parent_path();
        for (const auto& input : inputs) convert_one(input, output_dir, options);
    }
    std::cout << "Concluido: " << inputs.size() << " arquivo(s).\n";
    return 0;
}

#ifdef _WIN32
std::string wide_to_utf8(const wchar_t* value) {
    if (!value) return {};
    const int size = WideCharToMultiByte(CP_UTF8, 0, value, -1, nullptr, 0, nullptr, nullptr);
    if (size <= 0) throw std::runtime_error("falha ao converter argumento Unicode");
    std::string result(static_cast<std::size_t>(size), '\0');
    WideCharToMultiByte(CP_UTF8, 0, value, -1, result.data(), size, nullptr, nullptr);
    result.pop_back();
    return result;
}

#ifdef DBC_GUI
std::wstring utf8_to_wide(const std::string& value) {
    if (value.empty()) return {};
    const int size = MultiByteToWideChar(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), nullptr, 0);
    if (size <= 0) return L"Erro desconhecido";
    std::wstring result(static_cast<std::size_t>(size), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), result.data(), size);
    return result;
}

enum GuiId {
    ID_INPUT = 100,
    ID_INPUT_FILE = 101,
    ID_INPUT_FOLDER = 102,
    ID_OUTPUT = 103,
    ID_OUTPUT_FOLDER = 104,
    ID_CSV = 110,
    ID_XLSX = 111,
    ID_PARQUET = 112,
    ID_DBF = 113,
    ID_JSONL = 114,
    ID_RECURSIVE = 115,
    ID_CONVERT = 130,
    ID_OPEN_OUTPUT = 131,
    ID_PROGRESS = 132,
    ID_STATUS = 133,
    ID_LOG = 134
};

constexpr UINT WM_CONVERSION_DONE = WM_APP + 1;
HWND gui_input = nullptr;
HWND gui_output = nullptr;
HWND gui_convert = nullptr;
HWND gui_open_output = nullptr;
HWND gui_progress = nullptr;
HWND gui_status = nullptr;
HWND gui_log = nullptr;
bool gui_busy = false;
HFONT gui_font_body = nullptr;
HFONT gui_font_step = nullptr;
HFONT gui_font_button = nullptr;
HFONT gui_font_title = nullptr;
HFONT gui_font_subtitle = nullptr;
HBRUSH gui_white_brush = nullptr;
HBITMAP gui_solppe_logo = nullptr;

std::wstring control_text(HWND control) {
    const int length = GetWindowTextLengthW(control);
    std::wstring value(static_cast<std::size_t>(length) + 1, L'\0');
    GetWindowTextW(control, value.data(), length + 1);
    value.resize(static_cast<std::size_t>(length));
    return value;
}

void append_log(const std::wstring& message) {
    const int end = GetWindowTextLengthW(gui_log);
    SendMessageW(gui_log, EM_SETSEL, end, end);
    SendMessageW(gui_log, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(message.c_str()));
}

void set_input_path(const std::wstring& value) {
    if (value.empty()) return;
    SetWindowTextW(gui_input, value.c_str());
    const fs::path input(value);
    std::error_code ec;
    const bool directory = fs::is_directory(input, ec);
    const fs::path output = (directory ? input : input.parent_path()) / L"convertidos";
    SetWindowTextW(gui_output, output.c_str());
}

std::wstring browse_folder(HWND owner, const wchar_t* title) {
    BROWSEINFOW browse{};
    browse.hwndOwner = owner;
    browse.lpszTitle = title;
    browse.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
    PIDLIST_ABSOLUTE item = SHBrowseForFolderW(&browse);
    if (!item) return {};
    std::array<wchar_t, MAX_PATH> path{};
    const bool ok = SHGetPathFromIDListW(item, path.data()) != FALSE;
    CoTaskMemFree(item);
    return ok ? std::wstring(path.data()) : std::wstring();
}

void choose_input_file(HWND owner) {
    std::array<wchar_t, 32768> path{};
    OPENFILENAMEW dialog{};
    dialog.lStructSize = sizeof(dialog);
    dialog.hwndOwner = owner;
    dialog.lpstrFilter = L"Arquivos DBC do DATASUS (*.dbc)\0*.dbc\0Todos os arquivos (*.*)\0*.*\0\0";
    dialog.lpstrFile = path.data();
    dialog.nMaxFile = static_cast<DWORD>(path.size());
    dialog.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_EXPLORER;
    if (!GetOpenFileNameW(&dialog)) return;
    set_input_path(path.data());
}

void choose_input_folder(HWND owner) {
    const std::wstring path = browse_folder(owner, L"Selecione a pasta que contém os arquivos DBC");
    if (path.empty()) return;
    set_input_path(path);
}

struct GuiJob {
    std::wstring input;
    std::wstring output;
    std::vector<std::string> formats;
    std::string encoding;
    std::string delimiter;
    bool recursive = false;
};

struct GuiResult {
    bool success = false;
    std::wstring message;
    std::wstring output;
};

bool checked(HWND window, int id) {
    return SendDlgItemMessageW(window, id, BM_GETCHECK, 0, 0) == BST_CHECKED;
}

void begin_conversion(HWND window) {
    if (gui_busy) return;
    GuiJob job;
    job.input = control_text(gui_input);
    job.output = control_text(gui_output);
    if (job.input.empty() || job.output.empty()) {
        MessageBoxW(window, L"Informe a entrada DBC e a pasta de saída.", L"DBC Converter", MB_ICONWARNING);
        return;
    }
    if (checked(window, ID_CSV)) job.formats.emplace_back("csv");
    if (checked(window, ID_XLSX)) job.formats.emplace_back("xlsx");
    if (checked(window, ID_PARQUET)) job.formats.emplace_back("parquet");
    if (checked(window, ID_DBF)) job.formats.emplace_back("dbf");
    if (checked(window, ID_JSONL)) job.formats.emplace_back("jsonl");
    if (job.formats.empty()) {
        MessageBoxW(window, L"Selecione pelo menos um formato de saída.", L"DBC Converter", MB_ICONWARNING);
        return;
    }
    job.recursive = checked(window, ID_RECURSIVE);
    job.encoding = "cp1252";
    job.delimiter = "semicolon";

    gui_busy = true;
    EnableWindow(gui_convert, FALSE);
    EnableWindow(gui_open_output, FALSE);
    ShowWindow(gui_progress, SW_SHOW);
    SendMessageW(gui_progress, PBM_SETMARQUEE, TRUE, 25);
    SetWindowTextW(gui_status, L"Convertendo...");
    SetWindowTextW(gui_log, L"Iniciando conversão. Arquivos grandes podem levar alguns minutos.\r\n");

    std::thread([window, job = std::move(job)]() mutable {
        auto* result = new GuiResult;
        result->output = job.output;
        try {
            std::vector<std::string> args{"dbc_converter_gui", wide_to_utf8(job.input.c_str()), "-o", wide_to_utf8(job.output.c_str())};
            for (const auto& format : job.formats) {
                args.emplace_back("-f");
                args.push_back(format);
            }
            args.emplace_back("--encoding");
            args.push_back(job.encoding);
            args.emplace_back("--delimiter");
            args.push_back(job.delimiter);
            if (job.recursive) args.emplace_back("--recursive");
            const int code = run(args);
            result->success = code == 0;
            result->message = result->success ? L"Conversão concluída com sucesso." : L"A conversão terminou com erro.";
        } catch (const std::exception& error) {
            result->message = L"Erro: " + utf8_to_wide(error.what());
        }
        PostMessageW(window, WM_CONVERSION_DONE, 0, reinterpret_cast<LPARAM>(result));
    }).detach();
}

HWND make_control(HWND parent, const wchar_t* cls, const wchar_t* text, DWORD style,
                  int x, int y, int width, int height, int id) {
    HWND control = CreateWindowExW(cls == std::wstring(L"EDIT") ? WS_EX_CLIENTEDGE : 0,
                                   cls, text, WS_CHILD | WS_VISIBLE | style,
                                   x, y, width, height, parent,
                                   reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
                                   GetModuleHandleW(nullptr), nullptr);
    SendMessageW(control, WM_SETFONT,
                 reinterpret_cast<WPARAM>(gui_font_body ? gui_font_body : GetStockObject(DEFAULT_GUI_FONT)), TRUE);
    return control;
}

LRESULT CALLBACK gui_window_proc(HWND window, UINT message, WPARAM wparam, LPARAM lparam) {
    switch (message) {
        case WM_CREATE: {
            const HDC screen = GetDC(window);
            const int dpi = GetDeviceCaps(screen, LOGPIXELSY);
            ReleaseDC(window, screen);
            gui_font_body = CreateFontW(-MulDiv(10, dpi, 72), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                                        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                        CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
            gui_font_step = CreateFontW(-MulDiv(11, dpi, 72), 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE,
                                        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                        CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
            gui_font_button = CreateFontW(-MulDiv(11, dpi, 72), 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE,
                                          DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                          CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
            gui_font_title = CreateFontW(-MulDiv(20, dpi, 72), 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE,
                                         DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                         CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
            gui_font_subtitle = CreateFontW(-MulDiv(10, dpi, 72), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                                            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                            CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
            gui_white_brush = CreateSolidBrush(RGB(255, 255, 255));
            gui_solppe_logo = reinterpret_cast<HBITMAP>(LoadImageW(
                GetModuleHandleW(nullptr), MAKEINTRESOURCEW(200), IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR));
            DragAcceptFiles(window, TRUE);

            HWND step1 = make_control(window, L"STATIC", L"1. Escolha o arquivo DBC", 0, 22, 94, 300, 24, 0);
            SendMessageW(step1, WM_SETFONT, reinterpret_cast<WPARAM>(gui_font_step), TRUE);
            gui_input = make_control(window, L"EDIT", L"", ES_AUTOHSCROLL, 22, 120, 472, 29, ID_INPUT);
            make_control(window, L"BUTTON", L"Selecionar DBC...", BS_PUSHBUTTON, 504, 119, 126, 31, ID_INPUT_FILE);
            make_control(window, L"BUTTON", L"Pasta...", BS_PUSHBUTTON, 638, 119, 82, 31, ID_INPUT_FOLDER);
            HWND hint = make_control(window, L"STATIC", L"Também é possível arrastar um arquivo ou uma pasta para esta janela.",
                                     0, 24, 154, 620, 20, 0);
            SendMessageW(hint, WM_SETFONT, reinterpret_cast<WPARAM>(gui_font_body), TRUE);

            HWND step2 = make_control(window, L"STATIC", L"2. Confirme onde salvar", 0, 22, 184, 300, 24, 0);
            SendMessageW(step2, WM_SETFONT, reinterpret_cast<WPARAM>(gui_font_step), TRUE);
            gui_output = make_control(window, L"EDIT", L"", ES_AUTOHSCROLL, 22, 210, 608, 29, ID_OUTPUT);
            make_control(window, L"BUTTON", L"Alterar...", BS_PUSHBUTTON, 638, 209, 82, 31, ID_OUTPUT_FOLDER);
            make_control(window, L"STATIC", L"A pasta ‘convertidos’ é sugerida automaticamente e o DBC original não é alterado.",
                         0, 24, 244, 650, 20, 0);

            HWND step3 = make_control(window, L"STATIC", L"3. Escolha os formatos", 0, 22, 274, 300, 24, 0);
            SendMessageW(step3, WM_SETFONT, reinterpret_cast<WPARAM>(gui_font_step), TRUE);
            make_control(window, L"BUTTON", L"CSV", BS_AUTOCHECKBOX, 24, 302, 65, 24, ID_CSV);
            make_control(window, L"BUTTON", L"Excel (.xlsx)", BS_AUTOCHECKBOX, 108, 302, 112, 24, ID_XLSX);
            make_control(window, L"BUTTON", L"Parquet", BS_AUTOCHECKBOX, 242, 302, 88, 24, ID_PARQUET);
            make_control(window, L"BUTTON", L"DBF", BS_AUTOCHECKBOX, 350, 302, 62, 24, ID_DBF);
            make_control(window, L"BUTTON", L"JSONL", BS_AUTOCHECKBOX, 432, 302, 78, 24, ID_JSONL);
            make_control(window, L"BUTTON", L"Incluir subpastas", BS_AUTOCHECKBOX, 534, 302, 160, 24, ID_RECURSIVE);
            SendDlgItemMessageW(window, ID_CSV, BM_SETCHECK, BST_CHECKED, 0);
            SendDlgItemMessageW(window, ID_XLSX, BM_SETCHECK, BST_CHECKED, 0);
            SendDlgItemMessageW(window, ID_PARQUET, BM_SETCHECK, BST_CHECKED, 0);

            gui_convert = make_control(window, L"BUTTON", L"Converter arquivos", BS_OWNERDRAW, 22, 342, 224, 48, ID_CONVERT);
            SendMessageW(gui_convert, WM_SETFONT, reinterpret_cast<WPARAM>(gui_font_button), TRUE);
            gui_progress = CreateWindowExW(0, PROGRESS_CLASSW, nullptr, WS_CHILD | PBS_MARQUEE,
                                            266, 359, 454, 14, window,
                                            reinterpret_cast<HMENU>(ID_PROGRESS), GetModuleHandleW(nullptr), nullptr);
            gui_status = make_control(window, L"STATIC", L"Pronto para converter.", 0, 22, 405, 480, 22, ID_STATUS);
            gui_open_output = make_control(window, L"BUTTON", L"Abrir pasta de saída", BS_PUSHBUTTON, 568, 399, 152, 31, ID_OPEN_OUTPUT);
            EnableWindow(gui_open_output, FALSE);
            gui_log = make_control(window, L"EDIT", L"Selecione o DBC. A pasta de saída será preenchida automaticamente.",
                                   ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL | WS_VSCROLL,
                                   22, 438, 698, 105, ID_LOG);
            break;
        }
        case WM_PAINT: {
            PAINTSTRUCT paint{};
            HDC dc = BeginPaint(window, &paint);
            RECT client{};
            GetClientRect(window, &client);
            RECT header{0, 0, client.right, 76};
            FillRect(dc, &header, gui_white_brush);
            HBRUSH support_blue = CreateSolidBrush(RGB(220, 235, 251));
            RECT divider{0, 74, client.right, 76};
            FillRect(dc, &divider, support_blue);
            DeleteObject(support_blue);
            if (gui_solppe_logo) {
                BITMAP bitmap{};
                GetObjectW(gui_solppe_logo, sizeof(bitmap), &bitmap);
                HDC memory = CreateCompatibleDC(dc);
                HGDIOBJ old = SelectObject(memory, gui_solppe_logo);
                SetStretchBltMode(dc, HALFTONE);
                StretchBlt(dc, 22, 18, 180, 36, memory, 0, 0,
                           bitmap.bmWidth, bitmap.bmHeight, SRCCOPY);
                SelectObject(memory, old);
                DeleteDC(memory);
            }
            SetBkMode(dc, TRANSPARENT);
            SetTextColor(dc, RGB(6, 46, 111));
            SelectObject(dc, gui_font_title);
            const wchar_t* title = L"DBC para Excel, CSV e Parquet";
            TextOutW(dc, 230, 10, title, static_cast<int>(wcslen(title)));
            SetTextColor(dc, RGB(17, 24, 39));
            SelectObject(dc, gui_font_subtitle);
            const wchar_t* subtitle = L"Conversão simples e segura de arquivos DBC do DATASUS";
            TextOutW(dc, 232, 43, subtitle, static_cast<int>(wcslen(subtitle)));
            EndPaint(window, &paint);
            break;
        }
        case WM_CTLCOLORSTATIC: {
            HDC dc = reinterpret_cast<HDC>(wparam);
            SetBkMode(dc, TRANSPARENT);
            SetTextColor(dc, RGB(55, 65, 81));
            return reinterpret_cast<LRESULT>(gui_white_brush);
        }
        case WM_DRAWITEM: {
            auto* item = reinterpret_cast<DRAWITEMSTRUCT*>(lparam);
            if (item && item->CtlID == ID_CONVERT) {
                const bool disabled = (item->itemState & ODS_DISABLED) != 0;
                const COLORREF color = disabled ? RGB(156, 163, 175) : RGB(67, 169, 247);
                HBRUSH brush = CreateSolidBrush(color);
                HPEN pen = CreatePen(PS_SOLID, 1, color);
                HGDIOBJ old_brush = SelectObject(item->hDC, brush);
                HGDIOBJ old_pen = SelectObject(item->hDC, pen);
                RoundRect(item->hDC, item->rcItem.left, item->rcItem.top,
                          item->rcItem.right, item->rcItem.bottom, 10, 10);
                SelectObject(item->hDC, old_brush);
                SelectObject(item->hDC, old_pen);
                DeleteObject(brush);
                DeleteObject(pen);
                SetBkMode(item->hDC, TRANSPARENT);
                SetTextColor(item->hDC, RGB(255, 255, 255));
                SelectObject(item->hDC, gui_font_button);
                DrawTextW(item->hDC, L"Converter arquivos", -1, &item->rcItem,
                          DT_CENTER | DT_VCENTER | DT_SINGLELINE);
                if (item->itemState & ODS_FOCUS) DrawFocusRect(item->hDC, &item->rcItem);
                return TRUE;
            }
            break;
        }
        case WM_DROPFILES: {
            HDROP drop = reinterpret_cast<HDROP>(wparam);
            std::array<wchar_t, 32768> path{};
            if (DragQueryFileW(drop, 0, path.data(), static_cast<UINT>(path.size())) > 0) set_input_path(path.data());
            DragFinish(drop);
            break;
        }
        case WM_COMMAND: {
            switch (LOWORD(wparam)) {
                case ID_INPUT_FILE: choose_input_file(window); break;
                case ID_INPUT_FOLDER: choose_input_folder(window); break;
                case ID_OUTPUT_FOLDER: {
                    const std::wstring path = browse_folder(window, L"Selecione a pasta de saída");
                    if (!path.empty()) SetWindowTextW(gui_output, path.c_str());
                    break;
                }
                case ID_CONVERT: begin_conversion(window); break;
                case ID_OPEN_OUTPUT: {
                    const std::wstring path = control_text(gui_output);
                    if (!path.empty()) ShellExecuteW(window, L"open", path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
                    break;
                }
            }
            break;
        }
        case WM_CONVERSION_DONE: {
            std::unique_ptr<GuiResult> result(reinterpret_cast<GuiResult*>(lparam));
            gui_busy = false;
            EnableWindow(gui_convert, TRUE);
            SendMessageW(gui_progress, PBM_SETMARQUEE, FALSE, 0);
            ShowWindow(gui_progress, SW_HIDE);
            SetWindowTextW(gui_status, result->message.c_str());
            append_log(L"\r\n" + result->message + L"\r\nSaída: " + result->output + L"\r\n");
            EnableWindow(gui_open_output, result->success ? TRUE : FALSE);
            MessageBoxW(window, result->message.c_str(), L"DBC Converter", result->success ? MB_ICONINFORMATION : MB_ICONERROR);
            break;
        }
        case WM_CLOSE:
            if (gui_busy && MessageBoxW(window, L"Há uma conversão em andamento. Deseja fechar mesmo assim?",
                                        L"DBC Converter", MB_YESNO | MB_ICONWARNING) != IDYES) return 0;
            DestroyWindow(window);
            break;
        case WM_DESTROY:
            for (HFONT font : {gui_font_body, gui_font_step, gui_font_button, gui_font_title, gui_font_subtitle}) {
                if (font) DeleteObject(font);
            }
            if (gui_white_brush) DeleteObject(gui_white_brush);
            if (gui_solppe_logo) DeleteObject(gui_solppe_logo);
            PostQuitMessage(0);
            break;
        default:
            return DefWindowProcW(window, message, wparam, lparam);
    }
    return 0;
}
#endif
#endif

} // namespace

#if defined(_WIN32) && defined(DBC_GUI)
int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int show_command) {
    SetProcessDPIAware();
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    INITCOMMONCONTROLSEX common{sizeof(common), ICC_PROGRESS_CLASS | ICC_STANDARD_CLASSES};
    InitCommonControlsEx(&common);
    const wchar_t* class_name = L"DBCConverterWindow";
    WNDCLASSEXW window_class{};
    window_class.cbSize = sizeof(window_class);
    window_class.lpfnWndProc = gui_window_proc;
    window_class.hInstance = instance;
    window_class.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    window_class.hIcon = LoadIconW(instance, MAKEINTRESOURCEW(1));
    window_class.hIconSm = LoadIconW(instance, MAKEINTRESOURCEW(1));
    window_class.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    window_class.lpszClassName = class_name;
    RegisterClassExW(&window_class);
    const int window_width = 758;
    const int window_height = 600;
    const int window_x = std::max(0, (GetSystemMetrics(SM_CXSCREEN) - window_width) / 2);
    const int window_y = std::max(0, (GetSystemMetrics(SM_CYSCREEN) - window_height) / 2);
    HWND window = CreateWindowExW(WS_EX_CONTROLPARENT, class_name, L"DBC para Excel, CSV e Parquet - SOLPPE",
                                  WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
                                  window_x, window_y, window_width, window_height,
                                  nullptr, nullptr, instance, nullptr);
    if (!window) return 2;
    ShowWindow(window, show_command);
    UpdateWindow(window);
    MSG message{};
    while (GetMessageW(&message, nullptr, 0, 0) > 0) {
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }
    CoUninitialize();
    return static_cast<int>(message.wParam);
}
#elif defined(_WIN32)
int wmain(int argc, wchar_t** argv) {
    try {
        SetConsoleOutputCP(CP_UTF8);
        SetConsoleCP(CP_UTF8);
        std::vector<std::string> args;
        args.reserve(static_cast<std::size_t>(argc));
        for (int i = 0; i < argc; ++i) args.push_back(wide_to_utf8(argv[i]));
        return run(args);
    } catch (const std::exception& error) {
        std::cerr << "ERRO: " << error.what() << "\n";
        return 2;
    }
}
#else
int main(int argc, char** argv) {
    try {
        std::vector<std::string> args(argv, argv + argc);
        return run(args);
    } catch (const std::exception& error) {
        std::cerr << "ERRO: " << error.what() << "\n";
        return 2;
    }
}
#endif
