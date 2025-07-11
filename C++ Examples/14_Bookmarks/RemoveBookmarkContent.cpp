#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/RemoveBookmarkContent.docx";
	std::wstring inputFile = DATAPATH"/Bookmark.docx";

	//Load the document from disk.
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the bookmark by name.            
	intrusive_ptr<Bookmark> bookmark = document->GetBookmarks()->GetItem(L"Test");

	intrusive_ptr<Paragraph> para = Object::Dynamic_cast<Paragraph>(bookmark->GetBookmarkStart()->GetOwner());
	int startIndex = para->GetChildObjects()->IndexOf(bookmark->GetBookmarkStart());
	para = Object::Dynamic_cast<Paragraph>(bookmark->GetBookmarkEnd()->GetOwner());
	int endIndex = para->GetChildObjects()->IndexOf(bookmark->GetBookmarkEnd());

	//Remove the content object, and Start from next of BookmarkStart object, end up with previous of BookmarkEnd object. 
	//This method is only to remove the content of the bookmark.
	for (int i = startIndex + 1; i < endIndex; i++)
	{
		para->GetChildObjects()->RemoveAt(startIndex + 1);
	}

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}