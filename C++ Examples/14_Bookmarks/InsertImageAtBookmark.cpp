#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/InsertImageAtBookmark.docx";
	std::wstring inputFile = DATAPATH"/Bookmark.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Create an instance of BookmarksNavigator
	intrusive_ptr<BookmarksNavigator> bn = new BookmarksNavigator(doc);

	//Find a bookmark named Test
	bn->MoveToBookmark(L"Test", true, true);

	//Add a section
	intrusive_ptr<Section> section0 = doc->AddSection();

	//Add a paragraph for the section
	intrusive_ptr<Paragraph> paragraph = section0->AddParagraph();

	//Add a picture into the paragraph
	intrusive_ptr<DocPicture> picture = paragraph->AppendPicture(DATAPATH"/Word.png");
	//Add the paragraph at the position of bookmark
	bn->InsertParagraph(paragraph);

	//Remove the section0
	doc->GetSections()->Remove(section0);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();

}