#include "pch.h"

#include <fstream>
#include <locale>
#include <codecvt>

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/GetMergeFieldName.txt";
	std::wstring inputFile = DATAPATH"/MailMerge.doc";

	std::wstring stringBuilder;

	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get merge field name
	std::vector<LPCWSTR_S> fieldNames = document->GetMailMerge()->GetMergeFieldNames();

	stringBuilder.append(L"The document has " + std::to_wstring(fieldNames.size()) + L" merge fields.");
	stringBuilder.append(L" The below is the name of the merge field:\n");
	for (auto name : fieldNames)
	{
		stringBuilder.append(name);
		stringBuilder.append(L"\n");
	}

	std::wofstream write(outputFile);
	auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
	write.imbue(LocUtf8);
	write << stringBuilder;
	write.close();

	document->Close();
}