#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Styles.docx";

	//Initialize a document
	intrusive_ptr<Document> document = new Document();
	intrusive_ptr<Section> sec = document->AddSection();
	//Add default title style to document and modify
	intrusive_ptr<Style> titleStyle = document->AddStyle(BuiltinStyle::Title);

	titleStyle->GetCharacterFormat()->SetFontName(L"cambria");
	titleStyle->GetCharacterFormat()->SetFontSize(28);

	titleStyle->GetCharacterFormat()->SetTextColor(Color::FromArgb(42, 123, 136));
	//Judge if it is Paragraph Style and then set paragraph format
	if (Object::Dynamic_cast<ParagraphStyle>(titleStyle) != nullptr)
	{
		intrusive_ptr<ParagraphStyle> ps = Object::Dynamic_cast<ParagraphStyle>(titleStyle);
		ps->GetParagraphFormat()->GetBorders()->GetBottom()->SetBorderType(BorderStyle::Single);
		ps->GetParagraphFormat()->GetBorders()->GetBottom()->SetColor(Color::FromArgb(42, 123, 136));
		ps->GetParagraphFormat()->GetBorders()->GetBottom()->SetLineWidth(1.5f);
		ps->GetParagraphFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);
	}
	//Add default normal style and modify
	intrusive_ptr<Style> normalStyle = document->AddStyle(BuiltinStyle::Normal);

	normalStyle->GetCharacterFormat()->SetFontName(L"cambria");
	normalStyle->GetCharacterFormat()->SetFontSize(11);
	//Add default heading1 style
	intrusive_ptr<Style> heading1Style = document->AddStyle(BuiltinStyle::Heading1);

	heading1Style->GetCharacterFormat()->SetFontName(L"cambria");
	heading1Style->GetCharacterFormat()->SetFontSize(14);

	heading1Style->GetCharacterFormat()->SetBold(true);
	heading1Style->GetCharacterFormat()->SetTextColor(Color::FromArgb(42, 123, 136));
	//Add default heading2 style
	intrusive_ptr<Style> heading2Style = document->AddStyle(BuiltinStyle::Heading2);

	heading2Style->GetCharacterFormat()->SetFontName(L"cambria");
	heading2Style->GetCharacterFormat()->SetFontSize(12);

	heading2Style->GetCharacterFormat()->SetBold(true);

	//List style
	intrusive_ptr<ListStyle> bulletList = new ListStyle(document, ListType::Bulleted);

	bulletList->GetCharacterFormat()->SetFontName(L"cambria");
	bulletList->GetCharacterFormat()->SetFontSize(12);

	bulletList->SetName(L"bulletList");
	document->GetListStyles()->Add(bulletList);

	//Apply the style
	intrusive_ptr<Paragraph> paragraph = sec->AddParagraph();
	paragraph->AppendText(L"Your Name");
	paragraph->ApplyStyle(BuiltinStyle::Title);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"Address, City, ST ZIP Code | Telephone | Email");
	paragraph->ApplyStyle(BuiltinStyle::Normal);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"Objective");
	paragraph->ApplyStyle(BuiltinStyle::Heading1);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"To get started right away, just click any placeholder text (such as this) and start typing to replace it with your own.");
	paragraph->ApplyStyle(BuiltinStyle::Normal);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"Education");
	paragraph->ApplyStyle(BuiltinStyle::Heading1);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"DEGREE | DATE EARNED | SCHOOL");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"Major:Text");
	paragraph->GetListFormat()->ApplyStyle(L"bulletList");
	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"Minor:Text");
	paragraph->GetListFormat()->ApplyStyle(L"bulletList");
	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"Related coursework:Text");
	paragraph->GetListFormat()->ApplyStyle(L"bulletList");

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"Skills & Abilities");
	paragraph->ApplyStyle(BuiltinStyle::Heading1);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"MANAGEMENT");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"Think a document that looks this good has to be difficult to format? Think again! To easily apply any text formatting you see in this document with just a click, on the Home tab of the ribbon, check out Styles.");
	paragraph->GetListFormat()->ApplyStyle(L"bulletList");

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"COMMUNICATION");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"You delivered that big presentation to rave reviews. Don’t be shy about it now! This is the place to show how well you work and play with others.");
	paragraph->GetListFormat()->ApplyStyle(L"bulletList");

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"LEADERSHIP");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"Are you president of your fraternity, head of the condo board, or a team lead for your favorite charity? You’re a natural leader—tell it like it is!");
	paragraph->GetListFormat()->ApplyStyle(L"bulletList");

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"Experience");
	paragraph->ApplyStyle(BuiltinStyle::Heading1);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"JOB TITLE | COMPANY | DATES FROM - TO");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);

	paragraph = sec->AddParagraph();
	paragraph->AppendText(L"This is the place for a brief summary of your key responsibilities and most stellar accomplishments.");
	paragraph->GetListFormat()->ApplyStyle(L"bulletList");

	//Save to docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
