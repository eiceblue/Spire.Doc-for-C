#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Sample_UTF-7.txt";
	wstring outputFile = output_path + L"LoadTextWithEncoding.docx";

	intrusive_ptr<Document> document =  new Document();
	document->LoadText(inputFile.c_str(), Encoding::GetUTF7());
	//Save and launch the document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}
