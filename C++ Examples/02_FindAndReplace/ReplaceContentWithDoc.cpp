#include "pch.h"
#include <vector>
#include <regex>

using namespace Spire::Doc;
int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile_1 = input_path+L"ReplaceContentWithDoc.docx";
	wstring inputFile_2 = input_path + L"Insert.docx";
	wstring outputFile = output_path + L"ReplaceContentWithDoc.docx";

	//Create the first document
	intrusive_ptr<Document> document1 = new Document();

	//Load the first document from disk.
	document1->LoadFromFile(inputFile_1.c_str());

	//Create the second document
	intrusive_ptr<Document> document2 = new Document();

	//Load the second document from disk.
	document2->LoadFromFile(inputFile_2.c_str());

	//Get the first section of the first document 
	intrusive_ptr<Section> section1 = document1->GetSections()->GetItemInSectionCollection(0);

	//Create a regex
	intrusive_ptr<Regex> regex = new Regex(L"\\[MY_DOCUMENT\\]", RegexOptions::None);

	//Find the text by regex
	std::vector<intrusive_ptr<TextSelection>> textSections = document1->FindAllPattern(regex);

	//Travel the found strings
	for (intrusive_ptr<TextSelection> seletion : textSections)
	{

		//Get the para
		intrusive_ptr<Paragraph> para = seletion->GetAsOneRange()->GetOwnerParagraph();

		//Get textRange
		intrusive_ptr<TextRange> textRange = seletion->GetAsOneRange();

		//Get the para index
		int index = section1->GetBody()->GetChildObjects()->IndexOf(para);

		//Insert the paragraphs of document2
		for (int i = 0; i < document2->GetSections()->GetCount(); i++)
		{
			intrusive_ptr<Section> section2 = document2->GetSections()->GetItemInSectionCollection(i);
			for (int j = 0; j < section2->GetParagraphs()->GetCount(); j++)
			{
				intrusive_ptr<Paragraph> paragraph = section2->GetParagraphs()->GetItemInParagraphCollection(j);
				section1->GetBody()->GetChildObjects()->Insert(index, Object::Dynamic_cast<Paragraph>(paragraph->Clone()));
			}
		}
		//Remove the found textRange
		para->GetChildObjects()->Remove(textRange);
	}

	//Save the document.
	document1->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document1->Dispose();

}
