#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ASCIICharactersBulletStyle.docx";

	//Create a new document
	intrusive_ptr<Document> document = new Document();
	intrusive_ptr<Section> section = document->AddSection();

	//Create four list styles based on different ASCII characters
	intrusive_ptr<ListStyle> listStyle1 = new ListStyle(document, ListType::Bulleted);
	listStyle1->SetName(L"liststyle");
	listStyle1->GetLevels()->GetItem(0)->SetBulletCharacter(L"\x006e");
	listStyle1->GetLevels()->GetItem(0)->GetCharacterFormat()->SetFontName(L"Wingdings");
	document->GetListStyles()->Add(listStyle1);
	intrusive_ptr<ListStyle> listStyle2 = new ListStyle(document, ListType::Bulleted);
	listStyle2->SetName(L"liststyle2");
	listStyle2->GetLevels()->GetItem(0)->SetBulletCharacter(L"\x0075");
	listStyle2->GetLevels()->GetItem(0)->GetCharacterFormat()->SetFontName(L"Wingdings");
	document->GetListStyles()->Add(listStyle2);
	intrusive_ptr<ListStyle> listStyle3 = new ListStyle(document, ListType::Bulleted);
	listStyle3->SetName(L"liststyle3");
	listStyle3->GetLevels()->GetItem(0)->SetBulletCharacter(L"\x00b2");
	listStyle3->GetLevels()->GetItem(0)->GetCharacterFormat()->SetFontName(L"Wingdings");
	document->GetListStyles()->Add(listStyle3);
	intrusive_ptr<ListStyle> listStyle4 = new ListStyle(document, ListType::Bulleted);
	listStyle4->SetName(L"liststyle4");
	listStyle4->GetLevels()->GetItem(0)->SetBulletCharacter(L"\x00d8");
	listStyle4->GetLevels()->GetItem(0)->GetCharacterFormat()->SetFontName(L"Wingdings");
	document->GetListStyles()->Add(listStyle4);

	//Add four paragraphs and apply list style separately
	intrusive_ptr<Paragraph> p1 = section->GetBody()->AddParagraph();
	p1->AppendText(L"Spire.Doc for .NET");
	p1->GetListFormat()->ApplyStyle(listStyle1->GetName());
	p1->GetListFormat()->ApplyStyle(listStyle1->GetName());
	intrusive_ptr<Paragraph> p2 = section->GetBody()->AddParagraph();
	p2->AppendText(L"Spire.Doc for .NET");
	p2->GetListFormat()->ApplyStyle(listStyle2->GetName());
	intrusive_ptr<Paragraph> p3 = section->GetBody()->AddParagraph();
	p3->AppendText(L"Spire.Doc for .NET");
	p3->GetListFormat()->ApplyStyle(listStyle3->GetName());
	intrusive_ptr<Paragraph> p4 = section->GetBody()->AddParagraph();
	p4->AppendText(L"Spire.Doc for .NET");
	p4->GetListFormat()->ApplyStyle(listStyle4->GetName());

	//Save to docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
