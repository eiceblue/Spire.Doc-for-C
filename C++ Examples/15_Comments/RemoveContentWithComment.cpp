#include "pch.h"
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/RemoveContentWithComment.docx";
	std::wstring inputFile = DATAPATH"/Comments.docx";

	//Create a document
	intrusive_ptr<Document> document = new Document();

	//Load the document from disk.
	document->LoadFromFile(inputFile.c_str());

	//Get the first comment
	intrusive_ptr<Comment> comment = document->GetComments()->GetItem(0);

	//Get the paragraph of obtained comment
	intrusive_ptr<Paragraph> para = comment->GetOwnerParagraph();

	//Get index of the CommentMarkStart 
	int startIndex = para->GetChildObjects()->IndexOf(comment->GetCommentMarkStart());

	//Get index of the CommentMarkEnd
	int endIndex = para->GetChildObjects()->IndexOf(comment->GetCommentMarkEnd());

	//Create a list
	std::vector<intrusive_ptr<TextRange>> list;

	//Get TextRanges between the indexes
	for (int i = startIndex; i < endIndex; i++)
	{
		if (Object::CheckType<TextRange>(para->GetChildObjects()->GetItem(i)))
		{
			list.push_back(boost::dynamic_pointer_cast<TextRange>(para->GetChildObjects()->GetItem(i)));
		}
	}

	//Insert a new TextRange
	intrusive_ptr<TextRange> textRange = new TextRange(document);

	//Set text is null
	textRange->SetText(nullptr);

	//Insert the new textRange
	para->GetChildObjects()->Insert(endIndex, textRange);

	//Remove previous TextRanges
	for (size_t i = 0; i < list.size(); i++)
	{
		para->GetChildObjects()->Remove(list[i]);
	}

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}