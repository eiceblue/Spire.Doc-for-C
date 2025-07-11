#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/ReplaceBookmarkContent.docx";
	std::wstring inputFile = DATAPATH"/Bookmark.docx";

	//Load the document from disk.
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Locate the bookmark.
	intrusive_ptr<BookmarksNavigator> bookmarkNavigator = new BookmarksNavigator(doc);
	bookmarkNavigator->MoveToBookmark(L"Test");

	//Replace the context with new.
	bookmarkNavigator->ReplaceBookmarkContent(L"This is replaced content.", false);

	//Save the document.
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}