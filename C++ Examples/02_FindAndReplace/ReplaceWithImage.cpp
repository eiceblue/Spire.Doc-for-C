#include "pch.h"



using namespace Spire::Doc;

int main()
{
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile_1 = input_path + L"Template.docx";
	wstring inputFile_2 = input_path + L"E-iceblue.png";
	wstring outputFile = output_path + L"ReplaceWithImage.docx";
	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile_1.c_str());
	//Find the string "E-iceblue" in the document
	std::vector<intrusive_ptr<TextSelection>> selections = doc->FindAllString(L"E-iceblue", true, true);
	int index = 0;
	intrusive_ptr<TextRange> range = nullptr;

	//Remove the text and replace it with Image
	for (intrusive_ptr<TextSelection> selection : selections)
	{
		intrusive_ptr<DocPicture> pic = new DocPicture(doc);
		pic->LoadImageSpire(inputFile_2.c_str());

		range = selection->GetAsOneRange();
		index = range->GetOwnerParagraph()->GetChildObjects()->IndexOf(range);
		range->GetOwnerParagraph()->GetChildObjects()->Insert(index, pic);
		range->GetOwnerParagraph()->GetChildObjects()->Remove(range);
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();

}
