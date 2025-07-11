#include "pch.h"
using namespace Spire::Doc;


void CreateBookmark(intrusive_ptr<Section> section)
{
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();
	intrusive_ptr<TextRange> txtRange = paragraph->AppendText(L"The following example demonstrates how to create bookmark in a Word document.");
	txtRange->GetCharacterFormat()->SetItalic(true);

	section->AddParagraph();
	paragraph = section->AddParagraph();
	txtRange = paragraph->AppendText(L"Simple Create Bookmark.");
	txtRange->GetCharacterFormat()->SetTextColor(Color::GetCornflowerBlue());
	paragraph->ApplyStyle(BuiltinStyle::Heading2);

	//Write simple CreateBookmarks.
	section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph->AppendBookmarkStart(L"SimpleCreateBookmark");
	paragraph->AppendText(L"This is a simple bookmark.");
	paragraph->AppendBookmarkEnd(L"SimpleCreateBookmark");

	section->AddParagraph();
	paragraph = section->AddParagraph();
	txtRange = paragraph->AppendText(L"Nested Create Bookmark.");
	txtRange->GetCharacterFormat()->SetTextColor(Color::GetCornflowerBlue());
	paragraph->ApplyStyle(BuiltinStyle::Heading2);

	//Write nested CreateBookmarks.
	section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph->AppendBookmarkStart(L"Root");
	txtRange = paragraph->AppendText(L" This is Root data ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	paragraph->AppendBookmarkStart(L"NestedLevel1");
	txtRange = paragraph->AppendText(L" This is Nested Level1 ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	txtRange->GetCharacterFormat()->SetTextColor(Color::GetDarkSlateGray());
	paragraph->AppendBookmarkStart(L"NestedLevel2");
	txtRange = paragraph->AppendText(L" This is Nested Level2 ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	txtRange->GetCharacterFormat()->SetTextColor(Color::GetDimGray());
	paragraph->AppendBookmarkEnd(L"NestedLevel2");
	paragraph->AppendBookmarkEnd(L"NestedLevel1");
	paragraph->AppendBookmarkEnd(L"Root");

}
int main()
{
	std::wstring outputFile = OUTPUTPATH"/CreateBookmark.docx";

	//Create word document.
	intrusive_ptr<Document> document = new Document();

	//Create a new section.
	intrusive_ptr<Section> section = document->AddSection();

	CreateBookmark(section);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}