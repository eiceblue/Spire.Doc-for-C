#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"InputHtmlFile.html";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"HtmlFileToWord.docx";

	//Open an html file.
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str(), FileFormat::Html, XHTMLValidationType::None);

	//Save it to a Word document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}