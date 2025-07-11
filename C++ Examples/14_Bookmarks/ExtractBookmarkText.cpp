#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/ExtractBookmarkText.txt";
	std::wstring inputFile = DATAPATH"/BookmarkTemplate.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Creates a BookmarkNavigator instance to access the bookmark
	intrusive_ptr<BookmarksNavigator> navigator = new BookmarksNavigator(doc);
	//Locate a specific bookmark by bookmark name
	navigator->MoveToBookmark(L"Content");
	intrusive_ptr<TextBodyPart> textBodyPart = navigator->GetBookmarkContent();

	//Iterate through the items in the bookmark content to get the text
	std::wstring text = L"";
	for (int i = 0; i < textBodyPart->GetBodyItems()->GetCount(); i++)
	{
		intrusive_ptr<DocumentObject> item = textBodyPart->GetBodyItems()->GetItem(i);
		if (Object::CheckType<Paragraph>(item))
		{
			intrusive_ptr<Paragraph> paragraph = boost::dynamic_pointer_cast<Paragraph>(item);
			for (int j = 0; j < paragraph->GetChildObjects()->GetCount(); j++)
			{
				intrusive_ptr<DocumentObject> childObject = paragraph->GetChildObjects()->GetItem(j);
				if (Object::CheckType<TextRange>(childObject))
				{
					text += (boost::dynamic_pointer_cast<TextRange>(childObject))->GetText();
				}
			}
		}
	}

	//Save to TXT File and launch it
	std::wofstream foo(outputFile);
	foo << text;
	foo.close();
	doc->Close();
}