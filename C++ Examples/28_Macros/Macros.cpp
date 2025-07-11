#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Macros.docm";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Macros.docm";

	//Load the document
	intrusive_ptr<Document> document = new Document();
	
                //Loading documetn with macros.
	document->LoadFromFile(inputFile.c_str(), FileFormat::Docm);

	//Save docm file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docm);
	document->Close();
}
