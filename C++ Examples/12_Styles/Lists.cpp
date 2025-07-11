#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Lists.docx";

	//Initialize a document
	intrusive_ptr<Document> document = new Document();
	//Add a section
	intrusive_ptr<Section> sec = document->AddSection();
	//Add paragraph and set list style
	intrusive_ptr<Paragraph> paragraph = sec->AddParagraph();
	paragraph->AppendText(L"Lists");
	paragraph->ApplyStyle(BuiltinStyle::Title);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"Numbered List:")->GetCharacterFormat()->SetBold(true);

	//Create list style
	intrusive_ptr<ListStyle> numberList = new ListStyle(document, ListType::Numbered);
	numberList->SetName(L"numberList");
	//%1-%9
	numberList->GetLevels()->GetItem(1)->SetNumberPrefix(L"%1.");
	numberList->GetLevels()->GetItem(1)->SetPatternType(ListPatternType::Arabic);
	numberList->GetLevels()->GetItem(2)->SetNumberPrefix(L"%1.%2.");
	numberList->GetLevels()->GetItem(2)->SetPatternType(ListPatternType::Arabic);

	intrusive_ptr<ListStyle> bulletList = new ListStyle(document, ListType::Bulleted);
	bulletList->SetName(L"bulletList");

	//add the list style into document
	document->GetListStyles()->Add(numberList);
	document->GetListStyles()->Add(bulletList);

	//Add paragraph and apply the list style
	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"List Item 1");
	paragraph->GetListFormat()->ApplyStyle(numberList->GetName());

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"List Item 2");
	paragraph->GetListFormat()->ApplyStyle(numberList->GetName());

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"List Item 2.1");
	paragraph->GetListFormat()->ApplyStyle(numberList->GetName());
	paragraph->GetListFormat()->SetListLevelNumber(1);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"List Item 2.2");
	paragraph->GetListFormat()->ApplyStyle(numberList->GetName());
	paragraph->GetListFormat()->SetListLevelNumber(1);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"List Item 2.2.1");
	paragraph->GetListFormat()->ApplyStyle(numberList->GetName());
	paragraph->GetListFormat()->SetListLevelNumber(2);
	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"List Item 2.2.2");
	paragraph->GetListFormat()->ApplyStyle(numberList->GetName());
	paragraph->GetListFormat()->SetListLevelNumber(2);
	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"List Item 2.2.3");
	paragraph->GetListFormat()->ApplyStyle(numberList->GetName());
	paragraph->GetListFormat()->SetListLevelNumber(2);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"List Item 2.3");
	paragraph->GetListFormat()->ApplyStyle(numberList->GetName());
	paragraph->GetListFormat()->SetListLevelNumber(1);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"List Item 3");
	paragraph->GetListFormat()->ApplyStyle(numberList->GetName());

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"Bulleted List:")->GetCharacterFormat()->SetBold(true);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"List Item 1");
	paragraph->GetListFormat()->ApplyStyle(bulletList->GetName());
	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"List Item 2");
	paragraph->GetListFormat()->ApplyStyle(bulletList->GetName());

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"List Item 2.1");
	paragraph->GetListFormat()->ApplyStyle(bulletList->GetName());
	paragraph->GetListFormat()->SetListLevelNumber(1);
	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"List Item 2.2");
	paragraph->GetListFormat()->ApplyStyle(bulletList->GetName());
	paragraph->GetListFormat()->SetListLevelNumber(1);
	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"List Item 3");
	paragraph->GetListFormat()->ApplyStyle(bulletList->GetName());

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();

}