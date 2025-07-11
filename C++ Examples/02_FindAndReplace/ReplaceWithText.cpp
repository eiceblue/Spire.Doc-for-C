#include "pch.h"

using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Sample.docx";
	wstring outputFile = output_path + L"ReplaceWithText.docx";
	//Create word document
	intrusive_ptr<Document> document = new Document();

	//Load the document from disk.
	document->LoadFromFile(inputFile.c_str());

	//Replace text
	document->Replace(L"word", L"ReplacedText", false, true);

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}