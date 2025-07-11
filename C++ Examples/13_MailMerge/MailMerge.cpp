#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/MailMerge.doc";
	std::wstring inputFile = DATAPATH"/MailMerge.doc";

	//Create word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	std::vector<LPCWSTR_S> filedNames = { L"Contact Name", L"Fax", L"Date" };

	//C# TO C++ CONVERTER TODO TASK: There is no C++ equivalent to 'ToString':
	std::vector<LPCWSTR_S> filedValues = { L"John Smith", L"+1 (69) 123456", DateTime::GetNow()->GetDate()->ToString() };

	document->GetMailMerge()->Execute(filedNames, filedValues);

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Doc);
	document->Close();
}