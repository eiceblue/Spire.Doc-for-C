#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/InsertAddressBlockField.docx";
	std::wstring inputFile = DATAPATH"/SampleB_2.docx";


	//Open a Word document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	intrusive_ptr<Paragraph> par = section->AddParagraph();

	//Add address block in the paragraph
	intrusive_ptr<Field> field = par->AppendField(L"ADDRESSBLOCK", FieldType::FieldAddressBlock);

	//Set field code
	field->SetCode(L"ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\"");

	//Save to file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}