#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToEpub.doc";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ToEpub.epub";
	
	//Create a new document.
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());
	//Save the document to a Epub file.
	doc->SaveToFile(outputFile.c_str(), FileFormat::EPub);
	doc->Close();
}