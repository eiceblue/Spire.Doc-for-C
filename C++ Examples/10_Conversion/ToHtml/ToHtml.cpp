#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToHtmlTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ToHtml.html";

	//Create word document
	intrusive_ptr<Document>  document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Html);
	document->Close();
}