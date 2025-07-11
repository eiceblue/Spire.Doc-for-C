#include "pch.h"

using namespace Spire::Doc;

void InsertComments(intrusive_ptr<Section> section)
{
	//Insert comment.
	intrusive_ptr<Paragraph> paragraph = section->GetParagraphs()->GetItemInParagraphCollection(1);
	intrusive_ptr<Spire::Doc::Comment> comment = paragraph->AppendComment(L"Spire.Doc for .NET");
	comment->GetFormat()->SetAuthor(L"E-iceblue");
	comment->GetFormat()->SetInitial(L"CM");
}

int main()
{
	std::wstring outputFile = OUTPUTPATH"/Comment.docx";
	std::wstring inputFile = DATAPATH"/CommentTemplate.docx";

	//Load the document from disk.
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	InsertComments(document->GetSections()->GetItemInSectionCollection(0));

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}