#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_6.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetVariables.txt";

	RemoveDirectoryW(outputFile.c_str());

	intrusive_ptr<Document> document = new Document();
	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());
	wstring stringBuilder;

	stringBuilder.append(L"This document has following variables:\n");
	int variablesCount = document->GetVariables()->GetCount();
	for (int i = 0; i < variablesCount; i++)
	{
		std::wstring name = document->GetVariables()->GetNameByIndex(i);
		std::wstring value = document->GetVariables()->GetValueByIndex(i);
		stringBuilder.append(L"Name: " + name + L", " + L"Value: " + value);
		stringBuilder.append(L"\n");
	}
	//Save to file.
	wofstream out;
	out.open(outputFile);
	out.flush();
	out << stringBuilder.c_str();
	out.close();
	document->Close();

	//delete stringBuilder;
}
