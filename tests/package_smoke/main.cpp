#include "docxcpp/document.hpp"

int main() {
  docxcpp::Document document;
  document.add_heading("Package Smoke", 1);
  document.add_paragraph("installed package works");
  document.save("package-smoke.docx");
  return 0;
}
