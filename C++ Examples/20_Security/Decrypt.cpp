#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/Decrypt.docx";
	std::wstring inputFile = DATAPATH"/TemplateWithPassword.docx";

	//Create word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str(), FileFormat::Docx, L"E-iceblue");

	//Save as doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}