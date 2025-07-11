#include "pch.h"

using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"Comments.docx";
	wstring outputFile = output_path + L"ReplaceTextByRegex.docx";

	//Create a document
	intrusive_ptr<Document> sourceDoc = new Document();

	//Load the document from disk.
	sourceDoc->LoadFromFile(inputFile.c_str());

	//Create a destination document
	intrusive_ptr<Document> destinationDoc = new Document();

	//Add section for destination document
	intrusive_ptr<Section> destinationSec = destinationDoc->AddSection();

	//Get the first comment
	intrusive_ptr<Comment> comment = sourceDoc->GetComments()->GetItem(0);

	//Get the paragraph of obtained comment
	intrusive_ptr<Paragraph> para = comment->GetOwnerParagraph();

	//Get index of the CommentMarkStart 
	int startIndex = para->GetChildObjects()->IndexOf(comment->GetCommentMarkStart());

	//Get index of the CommentMarkEnd
	int endIndex = para->GetChildObjects()->IndexOf(comment->GetCommentMarkEnd());

	//Traverse paragraph ChildObjects
	for (int i = startIndex; i <= endIndex; i++)
	{
		//Clone the ChildObjects of source document
		intrusive_ptr<DocumentObject> doobj = para->GetChildObjects()->GetItem(i)->Clone();

		//Add to destination document 
		destinationSec->AddParagraph()->GetChildObjects()->Add(doobj);
	}
	//Save the destination document
	destinationDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	destinationDoc->Close();
}
