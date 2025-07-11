#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_RtfFile.rtf";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RtfToHtml.html";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Html);
	document->Close();

}