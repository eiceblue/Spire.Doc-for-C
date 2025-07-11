#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile_2003 = output_path + L"WordToWordML.xml";
	wstring outputFile_2007 = output_path + L"WordToWordXML.xml";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//For word 2003:
	document->SaveToFile(outputFile_2003.c_str(), FileFormat::WordML);

	//For word 2007:
	document->SaveToFile(outputFile_2007.c_str(), FileFormat::WordXml);
	document->Close();

}