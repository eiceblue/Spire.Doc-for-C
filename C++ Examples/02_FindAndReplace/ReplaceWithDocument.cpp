#include "pch.h"

using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile_1 = input_path + L"Text2.docx";
	wstring inputFile_2 = input_path + L"Text1.docx";
	wstring outputFile = output_path + L"ReplaceWithDocument.docx";
	//Load a template document 
	intrusive_ptr<Document> doc = new Document(inputFile_1.c_str());

	//Load another document to replace text
	intrusive_ptr<IDocument> replaceDoc = new Document(inputFile_2.c_str());
	//Replace specified text with the other document
	doc->Replace(L"Document1", replaceDoc, false, true);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
