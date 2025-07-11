#include "pch.h"
using namespace Spire::Doc;


int main() {
	
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Sample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddGutter.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Create a new section
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	//Set gutter
	section->GetPageSetup()->SetGutter(100.0f);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}