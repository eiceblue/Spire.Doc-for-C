#include "pch.h"


using namespace Spire::Doc;

int main()
{
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddTabStopsToParagraph.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Add a section.
	intrusive_ptr<Section> section = document->AddSection();

	//Add paragraph 1.
	intrusive_ptr<Paragraph> paragraph1 = section->AddParagraph();

	//Add tab and set its position (in points).
	intrusive_ptr<Tab> tab = paragraph1->GetFormat()->GetTabs()->AddTab(28);

	//Set tab alignment.
	tab->SetJustification(TabJustification::Left);

	//Move to next tab and append text.
	paragraph1->AppendText(L"\tWashing Machine");

	//Add another tab and set its position (in points).
	tab = paragraph1->GetFormat()->GetTabs()->AddTab(280);

	//Set tab alignment.
	tab->SetJustification(TabJustification::Left);

	//Specify tab leader type.
	tab->SetTabLeader(TabLeader::Dotted);

	//Move to next tab and append text.
	paragraph1->AppendText(L"\t$650");

	//Add paragraph 2.
	intrusive_ptr<Paragraph> paragraph2 = section->AddParagraph();

	//Add tab and set its position (in points).
	tab = paragraph2->GetFormat()->GetTabs()->AddTab(28);

	//Set tab alignment.
	tab->SetJustification(TabJustification::Left);

	//Move to next tab and append text.
	paragraph2->AppendText(L"\tRefrigerator");

	//Add another tab and set its position (in points).
	tab = paragraph2->GetFormat()->GetTabs()->AddTab(280);

	//Set tab alignment.
	tab->SetJustification(TabJustification::Left);

	//Specify tab leader type.
	tab->SetTabLeader(TabLeader::NoLeader);

	//Move to next tab and append text.
	paragraph2->AppendText(L"\t$800");

	//Save the Word document
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();

}
