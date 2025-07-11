#include "pch.h"
#include <locale>
#include <codecvt>

using namespace Spire::Doc;


int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ExtractText.docx";
	wstring outputFile = output_path + L"GetText.txt";

	//Load the document from disk.
	intrusive_ptr<Document> document =  new Document();
	document->LoadFromFile(inputFile.c_str());

	//get text from document
	wstring text = document->GetText();

	//create a new TXT File to save extracted text
	std::wofstream write(outputFile);
	auto LocUtf8 = locale(locale(""), new std::codecvt_utf8<wchar_t>);
	write.imbue(LocUtf8);
	write << text;
	write.close();
	document->Close();
}
