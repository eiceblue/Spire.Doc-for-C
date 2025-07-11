#include "pch.h"

using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Bookmark.docx";
	wstring outputFile = output_path + L"FromBookmark.docx";

	//Create the source document
	intrusive_ptr<Document> sourcedocument = new Document();

	//Load the source document from disk.
	sourcedocument->LoadFromFile(inputFile.c_str());

	//Create a destination document
	intrusive_ptr<Document> destinationDoc = new Document();

	//Add a section for destination document
	intrusive_ptr<Section> section = destinationDoc->AddSection();

	//Add a paragraph for destination document
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	//Locate the bookmark in source document
	intrusive_ptr<BookmarksNavigator> navigator = new BookmarksNavigator(sourcedocument);

	//Find bookmark by name
	navigator->MoveToBookmark(L"Test", true, true);

	//get text Body part
	intrusive_ptr<TextBodyPart> textBodyPart = navigator->GetBookmarkContent();

	//Create a TextRange type list
	std::vector<intrusive_ptr<TextRange>> list;

	//Traverse the items of text Body
	for (int i = 0; i < textBodyPart->GetBodyItems()->GetCount(); i++)
	{
		intrusive_ptr<DocumentObject> item = textBodyPart->GetBodyItems()->GetItem(i);
		//if it is paragraph
		if (Object::CheckType<Paragraph>(item))
		{
			//Traverse the ChildObjects of the paragraph
			for (int i = 0; i < (Object::Dynamic_cast<Paragraph>(item))->GetChildObjects()->GetCount(); i++)
			{
				intrusive_ptr<DocumentObject> childObject = (Object::Dynamic_cast<Paragraph>(item))->GetChildObjects()->GetItem(i);
				//if it is TextRange
				if (Object::CheckType<TextRange>(childObject))
				{
					//Add it into list
					intrusive_ptr<TextRange> range = boost::dynamic_pointer_cast<TextRange>(childObject);
					list.push_back(range);
				}
			}
		}
	}

	//Add the extract content to destinationDoc document
	for (size_t m = 0; m < list.size(); m++)
	{
		paragraph->GetChildObjects()->Add(list[m]->Clone());
	}

	//Save the document.
	destinationDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	destinationDoc->Close();

}
