#include "pch.h"
#include <algorithm>
#include <string>
#include <vector>

using namespace Spire::Doc;


void ReplaceWithHTML(intrusive_ptr<TextRange> textRange, std::vector<intrusive_ptr<DocumentObject>>& replacement)
{

	//textRange index
	int index = textRange->GetOwner()->GetChildObjects()->IndexOf(textRange);

	//get owener paragraph
	intrusive_ptr<Paragraph> paragraph = textRange->GetOwnerParagraph();

	//get owner text Body
	intrusive_ptr<Body> sectionBody = paragraph->GetOwnerTextBody();

	//get the index of paragraph in section
	int paragraphIndex = sectionBody->GetChildObjects()->IndexOf(paragraph);

	int replacementIndex = -1;
	if (index == 0)
	{
		//remove the first child object
		paragraph->GetChildObjects()->RemoveAt(0);

		replacementIndex = sectionBody->GetChildObjects()->IndexOf(paragraph);
	}
	else if (index == paragraph->GetChildObjects()->GetCount() - 1)
	{
		paragraph->GetChildObjects()->RemoveAt(index);
		replacementIndex = paragraphIndex + 1;
	}
	else
	{
		//split owner paragraph
		intrusive_ptr<Paragraph> paragraph1 = Object::Dynamic_cast<Paragraph>(paragraph->Clone());
		while (paragraph->GetChildObjects()->GetCount() > index)
		{
			paragraph->GetChildObjects()->RemoveAt(index);
		}
		int i = 0;
		int count = index + 1;
		while (i < count)
		{
			paragraph1->GetChildObjects()->RemoveAt(0);
			i += 1;
		}
		sectionBody->GetChildObjects()->Insert(paragraphIndex + 1, paragraph1);

		replacementIndex = paragraphIndex + 1;
	}

	//insert replacement
	int finalCount = replacement.size() - 1;
	for (int i = 0; i <= finalCount; i++)
	{
		sectionBody->GetChildObjects()->Insert(replacementIndex + i, replacement[i]->Clone());
	}
}

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"ReplaceWithHtml.docx";
	wifstream  dataHTML(input_path + L"InputHtml1.txt");
	wstring outputFile = output_path + L"ReplaceWithHtml.docx";

	wstring HTML(istreambuf_iterator<wchar_t>(dataHTML), {});


	//Load the document from disk.  
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//collect the objects which is used to replace text
	std::vector<intrusive_ptr<DocumentObject>> replacement;

	//create a temporary section
	intrusive_ptr<Section> tempSection = document->AddSection();

	//add a paragraph to append html
	intrusive_ptr<Paragraph> par = tempSection->AddParagraph();
	par->AppendHTML(HTML.c_str());

	//get the objects in temporary section
	for (int i = 0; i < tempSection->GetBody()->GetChildObjects()->GetCount(); i++)
	{
		intrusive_ptr<DocumentObject> obj = tempSection->GetBody()->GetChildObjects()->GetItem(i);
		intrusive_ptr<DocumentObject> docObj = Object::Dynamic_cast<DocumentObject>(obj);
		replacement.push_back(docObj);
	}

	//Find all text which will be replaced.
	std::vector<intrusive_ptr<TextSelection>> selections = document->FindAllString(L"[#placeholder]", false, true);

	std::vector<intrusive_ptr<TextRange>> locations;
	for (intrusive_ptr<TextSelection> selection : selections)
	{
		
		intrusive_ptr<TextRange> tempVar = selection->GetAsOneRange();
		locations.push_back(tempVar);
	}
	std::sort(locations.begin(), locations.end());

	for (intrusive_ptr<TextRange> location : locations)
	{
		//replace the text with HTML.c_str()
		ReplaceWithHTML(location, replacement);
	}

	//remove the temp section
	document->GetSections()->Remove(tempSection);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();

}

