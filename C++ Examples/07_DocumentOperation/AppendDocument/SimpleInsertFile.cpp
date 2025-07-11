#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile_1 = input_path + L"Template_N5.docx";
	wstring inputFile_2 = input_path + L"Template_N3.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SimpleInsertFile.docx";

	//Load the Word document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile_1.c_str());

	//Insert document from file
	doc->InsertTextFromFile(inputFile_2.c_str(), FileFormat::Auto);

	//Save the document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
}
