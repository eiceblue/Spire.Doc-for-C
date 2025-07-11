#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_HtmlFile.html";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"HtmlToPdf.pdf";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str(), FileFormat::Html, XHTMLValidationType::None);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	document->Close();
}
