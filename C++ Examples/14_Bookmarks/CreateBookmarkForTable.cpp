#include "pch.h"
using namespace Spire::Doc;

void CreateBookmarkForTable(intrusive_ptr<Document> doc, intrusive_ptr<Section> section)
{
	//Add a paragraph
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	//Append text for added paragraph
	intrusive_ptr<TextRange> txtRange = paragraph->AppendText(L"The following example demonstrates how to create bookmark for a table in a Word document.");

	//Set the font in italic
	txtRange->GetCharacterFormat()->SetItalic(true);

	//Append bookmark start
	paragraph->AppendBookmarkStart(L"CreateBookmark");

	//Append bookmark end
	paragraph->AppendBookmarkEnd(L"CreateBookmark");

	//Add table
	intrusive_ptr<Table> table = section->AddTable(true);

	//Set the number of rows and columns
	table->ResetCells(2, 2);

	//Append text for table cells		
	intrusive_ptr<TextRange> range = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"sampleA");
	range = table->GetRows()->GetItemInRowCollection(0)->GetCells()->GetItemInCellCollection(1)->AddParagraph()->AppendText(L"sampleB");
	range = table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(0)->AddParagraph()->AppendText(L"120");
	range = table->GetRows()->GetItemInRowCollection(1)->GetCells()->GetItemInCellCollection(1)->AddParagraph()->AppendText(L"260");

	//Get the bookmark by index.
	intrusive_ptr<Bookmark> bookmark = doc->GetBookmarks()->GetItem(0);

	//Get the name of bookmark.
	//std::wstring bookmarkName = bookmark->GetName();

	//Locate the bookmark by name.
	intrusive_ptr<BookmarksNavigator> navigator = new BookmarksNavigator(doc);
	navigator->MoveToBookmark(bookmark->GetName());

	//Add table to TextBodyPart
	intrusive_ptr<TextBodyPart> part = navigator->GetBookmarkContent();
	part->GetBodyItems()->Add(table);

	//Replace bookmark cotent with table
	navigator->ReplaceBookmarkContent(part);

	//delete navigator;
}

int main()
{
	std::wstring outputFile = OUTPUTPATH"/CreateBookmarkForTable.docx";
	//Create word document.
	intrusive_ptr<Document> document = new Document();

	//Add a new section.
	intrusive_ptr<Section> section = document->AddSection();

	//Create bookmark for a table
	CreateBookmarkForTable(document, section);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}