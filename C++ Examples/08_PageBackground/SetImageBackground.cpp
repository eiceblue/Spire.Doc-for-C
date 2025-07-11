#include "pch.h"
using namespace Spire::Doc;


int main() {

	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring inputFile_Img = input_path + L"Background.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetImageBackground.docx";

	//load a word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//set the background type as picture.
	document->GetBackground()->SetType(BackgroundType::Picture);

	//save the file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();

}