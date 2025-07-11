#include "pch.h"
#include <locale>
#include <codecvt>

using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring outputFile = output_path + L"CountWordsNumber.txt";

	//Create Word document.
	intrusive_ptr<Document> document =  new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Count the number of words.
	wstring content;
	content.append(L"CharCount: " + to_wstring(document->GetBuiltinDocumentProperties()->GetCharCount()));
	content.append(L"\n");
	content.append(L"CharCountWithSpace: " + to_wstring(document->GetBuiltinDocumentProperties()->GetCharCountWithSpace()));
	content.append(L"\n");
	content.append(L"WordCount: " + to_wstring(document->GetBuiltinDocumentProperties()->GetWordCount()));

	//Save to file.
	std::wofstream write(outputFile);
	auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
	write.imbue(LocUtf8);
	write << content;
	write.close();
	document->Close();
	document->Close();
}