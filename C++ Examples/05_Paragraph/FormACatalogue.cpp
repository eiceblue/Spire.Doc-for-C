#include "pch.h"


using namespace Spire::Doc;
using namespace std;
#define stringify(name) # name


int main()
{
	const wchar_t* convert_builtinStyleenum[] =
	{
		L"Heading1",
		L"Heading2",
		L"Heading3",
	};

	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"FormACatalogue.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Add a new section. 
	intrusive_ptr<Section> section = document->AddSection();
	intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetCount() > 0 ? section->GetParagraphs()->GetItemInParagraphCollection(0) : section->AddParagraph();

	//Add Heading 1.
	paragraph = section->AddParagraph();
	paragraph->AppendText(convert_builtinStyleenum[0]);
	paragraph->ApplyStyle(BuiltinStyle::Heading1);
	paragraph->GetListFormat()->ApplyNumberedStyle();

	//Add Heading 2.
	paragraph = section->AddParagraph();
	paragraph->AppendText(convert_builtinStyleenum[1]);
	paragraph->ApplyStyle(BuiltinStyle::Heading2);

	//List style for Headings 2.
	intrusive_ptr<ListStyle> listSty2 = new ListStyle(document, ListType::Numbered);

	for (int i = 0; i < listSty2->GetLevels()->GetCount(); i++)
	{
		intrusive_ptr<ListLevel> listLev = listSty2->GetLevels()->GetItem(i);
		listLev->SetUsePrevLevelPattern(true);
		listLev->SetNumberPrefix(L"1.");
	}
	listSty2->SetName(L"MyStyle2");
	document->GetListStyles()->Add(listSty2);
	paragraph->GetListFormat()->ApplyStyle(listSty2->GetName());

	//Add list style 3.
	intrusive_ptr<ListStyle> listSty3 = new ListStyle(document, ListType::Numbered);

	for (int i = 0; i < listSty3->GetLevels()->GetCount(); i++)
	{
		intrusive_ptr<ListLevel> listlev = listSty3->GetLevels()->GetItem(i);
		listlev->SetUsePrevLevelPattern(true);
		listlev->SetNumberPrefix(L"1.1.");
	}
	listSty3->SetName(L"MyStyle3");
	document->GetListStyles()->Add(listSty3);

	//Add Heading 3.
	for (int i = 0; i < 4; i++)
	{
		paragraph = section->AddParagraph();

		//Append text
		paragraph->AppendText(convert_builtinStyleenum[2]);

		//Apply list style 3 for Heading 3
		paragraph->ApplyStyle(BuiltinStyle::Heading3);
		paragraph->GetListFormat()->ApplyStyle(listSty3->GetName());
	}
	//Save the file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();

}
