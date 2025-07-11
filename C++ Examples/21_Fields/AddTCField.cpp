#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/AddTCField.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Add a new section.
	intrusive_ptr<Section> section = document->AddSection();

	//Add a new paragraph.
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	//Add TC field in the paragraph
	intrusive_ptr<Field> field = paragraph->AppendField(L"TC", FieldType::FieldTOCEntry);
	wstring codeString = L"TC ";
	codeString += L"\"Entry Text\"";
	codeString += L" \\f";
	codeString += L" t";
	field->SetCode(codeString.c_str());
	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}