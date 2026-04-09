if(NOT DEFINED INPUT)
  message(FATAL_ERROR "INPUT is required")
endif()

if(NOT DEFINED OUTPUT)
  message(FATAL_ERROR "OUTPUT is required")
endif()

file(READ "${INPUT}" ARCHIVE_HEX HEX)
string(LENGTH "${ARCHIVE_HEX}" ARCHIVE_HEX_LEN)

if(ARCHIVE_HEX_LEN EQUAL 0)
  message(FATAL_ERROR "Input archive is empty: ${INPUT}")
endif()

set(ARCHIVE_ROWS "")
math(EXPR LAST_INDEX "${ARCHIVE_HEX_LEN} - 2")
set(BYTES_IN_ROW 0)
set(CURRENT_ROW "")

foreach(INDEX RANGE 0 ${LAST_INDEX} 2)
  string(SUBSTRING "${ARCHIVE_HEX}" ${INDEX} 2 BYTE_HEX)
  string(APPEND CURRENT_ROW "0x${BYTE_HEX}, ")
  math(EXPR BYTES_IN_ROW "${BYTES_IN_ROW} + 1")

  if(BYTES_IN_ROW EQUAL 12)
    string(APPEND ARCHIVE_ROWS "    ${CURRENT_ROW}\n")
    set(CURRENT_ROW "")
    set(BYTES_IN_ROW 0)
  endif()
endforeach()

if(NOT CURRENT_ROW STREQUAL "")
  string(APPEND ARCHIVE_ROWS "    ${CURRENT_ROW}\n")
endif()

file(
  WRITE "${OUTPUT}"
  "#include <cstdint>\n"
  "#include <vector>\n"
  "\n"
  "#include \"docxcpp/opc_package.hpp\"\n"
  "#include \"internal/default_template_support.hpp\"\n"
  "\n"
  "namespace docxcpp {\n"
  "\n"
  "OpcPackage load_default_template_package() {\n"
  "  static const std::uint8_t kDefaultTemplateArchive[] = {\n"
  "${ARCHIVE_ROWS}"
  "  };\n"
  "  return OpcPackage::from_archive_bytes(std::vector<std::uint8_t>(\n"
  "      kDefaultTemplateArchive,\n"
  "      kDefaultTemplateArchive + sizeof(kDefaultTemplateArchive)));\n"
  "}\n"
  "\n"
  "}  // namespace docxcpp\n"
)
