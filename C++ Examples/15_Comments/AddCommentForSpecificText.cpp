#include "pch.h"

using namespace Spire::Doc;

void InsertComments(intrusive_ptr<Document> doc, const std::wstring& keystring)
{
	//Find the key string
	intrusive_ptr<TextSelection> find = doc->FindString(keystring.c_str(), false, true);

	//Create the commentmarkStart and commentmarkEnd
	intrusive_ptr<CommentMark> commentmarkStart = new CommentMark(doc);
	commentmarkStart->SetType(CommentMarkType::CommentStart);
	intrusive_ptr<CommentMark> commentmarkEnd = new CommentMark(doc);
	commentmarkEnd->SetType(CommentMarkType::CommentEnd);

	//Add the content for comment
	intrusive_ptr<Comment> comment = new Comment(doc);
	comment->GetBody()->AddParagraph()->SetText(L"Test comments");
	comment->GetFormat()->SetAuthor(L"E-iceblue");

	//Get the textRange
	intrusive_ptr<TextRange> range = find->GetAsOneRange();

	//Get its paragraph
	intrusive_ptr<Paragraph> para = range->GetOwnerParagraph();

	//Get the index of textRange 
	int index = para->GetChildObjects()->IndexOf(range);

	//Add comment
	para->GetChildObjects()->Add(comment);

	//Insert the commentmarkStart and commentmarkEnd
	para->GetChildObjects()->Insert(index, commentmarkStart);
	para->GetChildObjects()->Insert(index + 2, commentmarkEnd);
}

int main()
{
	std::wstring outputFile = OUTPUTPATH"/AddCommentForSpecificText.docx";
	std::wstring inputFile = DATAPATH"/CommentTemplate.docx";

	//Load the document from disk.
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	InsertComments(document, L"development");

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}