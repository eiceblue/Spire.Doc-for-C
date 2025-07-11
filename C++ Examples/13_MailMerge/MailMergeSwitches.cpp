#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/MailMergeSwitches.docx";
	std::wstring inputFile = DATAPATH"/MailMergeSwitches.docx";

	intrusive_ptr<Document> doc = new Document();
	//Load a mail merge template file
	doc->LoadFromFile(inputFile.c_str());

	std::vector<LPCWSTR_S> fieldName = {L"XX_Name"};
	std::vector<LPCWSTR_S> fieldValue = {L"Jason Tang"};

	doc->GetMailMerge()->Execute(fieldName, fieldValue);
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}