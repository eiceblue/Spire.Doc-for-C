#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/SpecifiedProtectionType.docx";
	std::wstring inputFile = DATAPATH"/Template_Docx_2.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Protect the Word file.
	document->Protect(ProtectionType::AllowOnlyReading, L"123456");

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}