#include "pch.h"


using namespace Spire::Doc;


int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"SetSpaceBetweenAsianAndLatinText.docx";
	wstring outputFile = output_path + L"SetSpaceBetweenAsianAndLatinText.docx";

	intrusive_ptr<Document> document = new Document();
	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	intrusive_ptr<Paragraph> para = document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0);

	//Set whether to automatically adjust space between Asian text and Latin text
	para->GetFormat()->SetAutoSpaceDE(false);
	//Set whether to automatically adjust space between Asian text and numbers
	para->GetFormat()->SetAutoSpaceDN(true);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();

}