#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/ExtractComment.txt";
	std::wstring inputFile = DATAPATH"/CommentSample.docx";

	//Load the document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	std::wstring stringBuilder;

	//Traverse all comments
	for (int i = 0; i < doc->GetComments()->GetCount(); i++)
	{
		intrusive_ptr<Comment> comment = doc->GetComments()->GetItem(i);
		for (int j = 0; j < comment->GetBody()->GetParagraphs()->GetCount(); j++)
		{
			intrusive_ptr<Paragraph> p = comment->GetBody()->GetParagraphs()->GetItemInParagraphCollection(j);
			stringBuilder.append(p->GetText());
			stringBuilder.append(L"\n");
		}
	}

	std::wofstream foo(outputFile);
	foo << stringBuilder.c_str();
	foo.close();
	doc->Close();

}