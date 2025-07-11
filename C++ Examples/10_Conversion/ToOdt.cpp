#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToOdt.doc";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ToOdt.odt";

	//Create word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Odt);
	document->Close();
}