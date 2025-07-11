#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/GetBookmarks.txt";
	std::wstring inputFile = DATAPATH"/Bookmarks.docx";

	//Create word document
	//Load the document from disk.
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the bookmark by index.
	intrusive_ptr<Bookmark> bookmark1 = document->GetBookmarks()->GetItem(0);

	//Get the bookmark by name.
	intrusive_ptr<Bookmark> bookmark2 = document->GetBookmarks()->GetItem(L"Test2");

	//Create StringBuilder to save 
	std::wstring content;

	//Set string format for displaying
	content.append(L"The bookmark obtained by index is ");
	content.append(bookmark1->GetName());
	content.append(L".\n");
	content.append(L"The bookmark obtained by name is ");
	content.append(bookmark2->GetName());
	content.append(L".\n");

	//Save them to a txt file
	std::wofstream foo(outputFile);
	foo << content.c_str();
	foo.close();
	document->Close();
}