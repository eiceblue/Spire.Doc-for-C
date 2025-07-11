#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile =OUTPUTPATH"/RemoveAndReplaceComment.docx";
	std::wstring inputFile = DATAPATH"/CommentSample.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Replace the content of the first comment
	doc->GetComments()->GetItem(0)->GetBody()->GetParagraphs()->GetItemInParagraphCollection(0)->Replace(L"This is the title", L"This comment is changed.", false, false);

	//Remove the second comment
	doc->GetComments()->RemoveAt(1);

	//Save and launch
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}