#include "pch.h"
using namespace Spire::Doc;


int main()
{
	std::wstring outputFile = OUTPUTPATH"/RemoveBookmark.docx";
	std::wstring inputFile = DATAPATH"/Bookmark.docx";

	//Load the document from disk.
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the bookmark by name.
	intrusive_ptr<Bookmark> bookmark = document->GetBookmarks()->GetItem(L"Test");

	//Remove the bookmark, not its content.
	document->GetBookmarks()->Remove(bookmark);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
