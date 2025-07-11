#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertPageBreakFirstApproach.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Find the specified word "technology" where we want to insert the page break.
	std::vector<intrusive_ptr<TextSelection>> selections = document->FindAllString(L"technology", true, true);

	//Traverse each word "technology".
	for (intrusive_ptr<TextSelection> ts : selections)
	{
		intrusive_ptr<TextRange> range = ts->GetAsOneRange();
		intrusive_ptr<Paragraph> paragraph = range->GetOwnerParagraph();
		int index = paragraph->GetChildObjects()->IndexOf(range);

		//Create a new instance of page break and insert a page break after the word "technology".
		intrusive_ptr<Break> pageBreak = new Break(document, BreakType::PageBreak);
		paragraph->GetChildObjects()->Insert(index + 1, pageBreak);

		//C# TO C++ CONVERTER TODO TASK: A 'delete pageBreak' statement was not added since pageBreak was passed to a method or constructor. Handle memory management manually.
	}

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();



}
