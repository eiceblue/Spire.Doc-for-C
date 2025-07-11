#include "pch.h"

using namespace Spire::Doc;

int main() 
{
	std::wstring outputFile = OUTPUTPATH"/InsertDocAtBookmark.docx";
	std::wstring inputFile = DATAPATH"/Bookmark.docx";

	//Create the first document
	intrusive_ptr<Document> document1 = new Document();

	//Load the first document from disk.
	document1->LoadFromFile(inputFile.c_str());

	//Create the second document
	intrusive_ptr<Document> document2 = new Document();

	//Load the second document from disk.
	document2->LoadFromFile(DATAPATH"/Insert.docx");

	//Get the first section of the first document 
	intrusive_ptr<Section> section1 = document1->GetSections()->GetItemInSectionCollection(0);

	//Locate the bookmark
	intrusive_ptr<BookmarksNavigator> bn = new BookmarksNavigator(document1);

	//Find bookmark by name
	bn->MoveToBookmark(L"Test", true, true);

	//Get bookmarkStart
	intrusive_ptr<BookmarkStart> start = bn->GetCurrentBookmark()->GetBookmarkStart();

	//Get the owner paragraph
	intrusive_ptr<Paragraph> para = start->GetOwnerParagraph();

	//Get the para index
	int index = section1->GetBody()->GetChildObjects()->IndexOf(para);

	//Insert the paragraphs of document2
	for (int i = 0; i < document2->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section2 = document2->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < section2->GetParagraphs()->GetCount(); j++)
		{
			intrusive_ptr<Paragraph> paragraph = section2->GetParagraphs()->GetItemInParagraphCollection(j);
			section1->GetBody()->GetChildObjects()->Insert(index + 1, Object::Dynamic_cast<Paragraph>(paragraph->Clone()));
		}
	}

	//Save the document.
	document1->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document1->Close();
}