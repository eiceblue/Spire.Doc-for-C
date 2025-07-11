#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConvertedTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ToPostScript.ps";

	//Create Word document.
	intrusive_ptr<Document> doc = new Document();

	//Load the file from disk.
	doc->LoadFromFile(inputFile.c_str());
	//Save to PCL file.
	doc->SaveToFile(outputFile.c_str(), FileFormat::PostScript);
	doc->Close();
}
