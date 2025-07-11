#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Sample.docx";
	wstring outputFile = output_path + L"InsertNewText.docx";

	//Load Document   
	intrusive_ptr<Document> doc =  new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Find all the text string "New Zealand” from the sample document
	std::vector<intrusive_ptr<TextSelection>> selections = doc->FindAllString(L"Word", true, true);
	int index = 0;

	//Defines text range
	intrusive_ptr<TextRange> range = new TextRange(doc);

	//Insert new text string (New) after the searched text string
	for (intrusive_ptr<TextSelection> selection : selections)
	{
		range = selection->GetAsOneRange();
		intrusive_ptr<TextRange> newrange = new TextRange(doc);
		newrange->SetText(L"New text)");
		index = range->GetOwnerParagraph()->GetChildObjects()->IndexOf(range);
		range->GetOwnerParagraph()->GetChildObjects()->Insert(index + 1, newrange);
	}

	//Find and highlight the newly added text string New
	std::vector<intrusive_ptr<TextSelection>> text = doc->FindAllString(L"New text", true, true);
	for (intrusive_ptr<TextSelection> seletion : text)
	{
		seletion->GetAsOneRange()->GetCharacterFormat()->SetHighlightColor(Color::GetYellow());
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}
