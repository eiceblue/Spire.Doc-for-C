#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/Encrypt.docx";
	std::wstring inputFile = DATAPATH"/Template.docx";

	//Create word document
	intrusive_ptr<Document> document = new Document();

	//Load Word document.
	document->LoadFromFile(inputFile.c_str());

	//encrypt document with password specified by textBox1
	document->Encrypt(L"E-iceblue");

	//Save as docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}