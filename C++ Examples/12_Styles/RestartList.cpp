#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RestartList.docx";
	//Create word document
	intrusive_ptr<Document> document = new Document();

	//Create a new section
	intrusive_ptr<Section> section = document->AddSection();

	//Create a new paragraph
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	//Append Text
	paragraph->AppendText(L"List 1");

	intrusive_ptr<ListStyle> numberList = new ListStyle(document, ListType::Numbered);
	numberList->SetName(L"Numbered1");
	document->GetListStyles()->Add(numberList);

	//Add paragraph and apply the list style
	paragraph = section->AddParagraph();
	paragraph->AppendText(L"List Item 1");
	paragraph->GetListFormat()->ApplyStyle(numberList->GetName());

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"List Item 2");
	paragraph->GetListFormat()->ApplyStyle(numberList->GetName());

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"List Item 3");
	paragraph->GetListFormat()->ApplyStyle(numberList->GetName());

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"List Item 4");
	paragraph->GetListFormat()->ApplyStyle(numberList->GetName());

	//Append Text
	paragraph = section->AddParagraph();
	paragraph->AppendText(L"List 2");

	intrusive_ptr<ListStyle> numberList2 = new ListStyle(document, ListType::Numbered);
	numberList2->SetName(L"Numbered2");
	//set start number of second list
	numberList2->GetLevels()->GetItem(0)->SetStartAt(10);
	document->GetListStyles()->Add(numberList2);

	//Add paragraph and apply the list style
	paragraph = section->AddParagraph();
	paragraph->AppendText(L"List Item 5");
	paragraph->GetListFormat()->ApplyStyle(numberList2->GetName());

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"List Item 6");
	paragraph->GetListFormat()->ApplyStyle(numberList2->GetName());

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"List Item 7");
	paragraph->GetListFormat()->ApplyStyle(numberList2->GetName());

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"List Item 8");
	paragraph->GetListFormat()->ApplyStyle(numberList2->GetName());

	//Save to docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
