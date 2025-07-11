#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"DetectAndRemoveVBAMacros.docm";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DetectAndRemoveVBAMacros.docm";

	//Load the document
	intrusive_ptr<Document> document = new Document();
	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//If the document contains Macros, remove them from the document.
	if (document->GetIsContainMacro())
	{
		document->ClearMacros();
	}

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docm);
	document->Close();
}
