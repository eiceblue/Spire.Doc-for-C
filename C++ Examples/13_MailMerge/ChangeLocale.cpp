#include "pch.h"
#include <locale>
#include <ctime>

using namespace Spire::Doc;

int main() 
{
	std::wstring outputFile = OUTPUTPATH"/ChangeLocale.doc";
	std::wstring inputFile = DATAPATH"/MailMerge.doc";

	//Load word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());
	// Store the current culture so it can be set back once mail merge is complete.
	std::locale originalLocale = std::locale::global(std::locale(""));

	std::locale germanLocale("de_DE.utf8");
	std::locale::global(germanLocale);

	std::time_t currentTime = std::time(nullptr);
	std::tm* localTime = std::localtime(&currentTime);

	std::wstring timeStr;
	timeStr.resize(100);
	std::wcsftime(&timeStr[0], timeStr.size(), L"%c", localTime);

	std::vector<LPCWSTR_S> fieldNames = { L"Contact Name", L"Fax", L"Date" };
	std::vector<LPCWSTR_S> fieldValues = { L"John Smith", L"+1 (69) 123456", timeStr.c_str() };
	document->GetMailMerge()->Execute(fieldNames, fieldValues);

	std::locale::global(originalLocale);
	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Doc);
	document->Close();
}