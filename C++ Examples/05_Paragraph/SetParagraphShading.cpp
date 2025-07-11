#include "pch.h"


using namespace Spire::Doc;


int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Template_Docx_1.docx";
	wstring outputFile = output_path + L"SetParagraphShading.docx";
	//Create Word document.
	intrusive_ptr<Document> document =  new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());
	//Get a paragraph.
	intrusive_ptr<Paragraph> paragaph = document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(0);

	//Set background color for the paragraph.
	paragaph->GetFormat()->SetBackColor(Color::GetYellow());

	//Set background color for the selected text of paragraph.
	paragaph = document->GetSections()->GetItemInSectionCollection(0)->GetParagraphs()->GetItemInParagraphCollection(2);
	intrusive_ptr<TextSelection> selection = paragaph->Find(L"Christmas", true, false);
	intrusive_ptr<TextRange> range = selection->GetAsOneRange();
	range->GetCharacterFormat()->SetTextBackgroundColor(Color::GetYellow());

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}