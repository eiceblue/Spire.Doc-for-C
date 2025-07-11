#include "pch.h"
#include <vector>

using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ReplaceTextWithField.docx";
	wstring outputFile = output_path + L"ReplaceTextWithField.docx";
	//Create word document
	intrusive_ptr<Document> document = new Document();

	//Load the document from disk.
	document->LoadFromFile(inputFile.c_str());

	//Find the target text
	intrusive_ptr<TextSelection> selection = document->FindString(L"summary", false, true);
	//Get text range
	intrusive_ptr<TextRange> textRange = selection->GetAsOneRange();
	//Get it's owner paragraph
	intrusive_ptr<Paragraph> ownParagraph = textRange->GetOwnerParagraph();
	//Get the index of this text range
	int rangeIndex = ownParagraph->GetChildObjects()->IndexOf(textRange);
	//Remove the text range
	ownParagraph->GetChildObjects()->RemoveAt(rangeIndex);
	//Remove the objects which are behind the text range
	std::vector<intrusive_ptr<DocumentObject>> tempList;
	for (int i = rangeIndex; i < ownParagraph->GetChildObjects()->GetCount(); i++)
	{
		//Add a copy of these objects into a temp list
		tempList.push_back(ownParagraph->GetChildObjects()->GetItem(rangeIndex)->Clone());
		ownParagraph->GetChildObjects()->RemoveAt(rangeIndex);
	}
	//Append field to the paragraph
	ownParagraph->AppendField(L"MyFieldName", FieldType::FieldMergeField);
	//Put these objects back into the paragraph one by one
	for (intrusive_ptr<DocumentObject> obj : tempList)
	{
		ownParagraph->GetChildObjects()->Add(obj);
	}

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
