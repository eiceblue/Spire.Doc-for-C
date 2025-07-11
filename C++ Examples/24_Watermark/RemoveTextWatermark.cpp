#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"RemoveTextWatermark.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveTextWatermark.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Set the watermark as null to remove the text and image watermark.
	document->SetWatermark(nullptr);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}

