#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_OdtFile.odt";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"OdtToWord.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Save to Docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();

}