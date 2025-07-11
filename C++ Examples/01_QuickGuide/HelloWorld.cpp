#include "pch.h"

using namespace Spire::Doc;

int main()
{
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"HelloWorld.docx";

	//Create word document
	intrusive_ptr<Document> document = new Document();

	//Create a new section
	intrusive_ptr<Section> section = document->AddSection();

	//Create a new paragraph
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	//Append Text
	paragraph->AppendText(L"Hello World!");

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);

	document->Close();
}