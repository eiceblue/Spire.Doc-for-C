#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/CopyBookmarkContent.docx";
	std::wstring inputFile = DATAPATH"/Bookmark.docx";

	//Load the document from disk.
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the bookmark by name.
	intrusive_ptr<Bookmark> bookmark = doc->GetBookmarks()->GetItem(L"Test");
	intrusive_ptr<DocumentObject> docObj = nullptr;

	//Judge if the paragraph includes the bookmark exists in the table, if it exists in cell,
	//Then need to find its outermost parent object(Table),
	//and get the start/end index of current object on GetBody().
	if ((Object::Dynamic_cast<Paragraph>(bookmark->GetBookmarkStart()->GetOwner()))->GetIsInCell())
	{
		docObj = bookmark->GetBookmarkStart()->GetOwner()->GetOwner()->GetOwner()->GetOwner();
	}
	else
	{
		docObj = bookmark->GetBookmarkStart()->GetOwner();
	}
	int startIndex = doc->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->IndexOf(docObj);
	if ((Object::Dynamic_cast<Paragraph>(bookmark->GetBookmarkEnd()->GetOwner()))->GetIsInCell())
	{
		docObj = bookmark->GetBookmarkEnd()->GetOwner()->GetOwner()->GetOwner()->GetOwner();
	}
	else
	{
		docObj = bookmark->GetBookmarkEnd()->GetOwner();
	}
	int endIndex = doc->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->IndexOf(docObj);

	//Get the start/end index of the bookmark object on the paragraph.
	intrusive_ptr<Paragraph> para = Object::Dynamic_cast<Paragraph>(bookmark->GetBookmarkStart()->GetOwner());
	int pStartIndex = para->GetChildObjects()->IndexOf(bookmark->GetBookmarkStart());
	para = Object::Dynamic_cast<Paragraph>(bookmark->GetBookmarkEnd()->GetOwner());
	int pEndIndex = para->GetChildObjects()->IndexOf(bookmark->GetBookmarkEnd());

	//Get the content of current bookmark and copy.
	intrusive_ptr<TextBodySelection> select = new TextBodySelection(doc->GetSections()->GetItemInSectionCollection(0)->GetBody(), startIndex, endIndex, pStartIndex, pEndIndex);
	intrusive_ptr<TextBodyPart> body = new TextBodyPart(select);
	for (int i = 0; i < body->GetBodyItems()->GetCount(); i++)
	{
		doc->GetSections()->GetItemInSectionCollection(0)->GetBody()->GetChildObjects()->Add((body->GetBodyItems())->GetItem(i)->Clone());
	}

	//Save the document.
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}