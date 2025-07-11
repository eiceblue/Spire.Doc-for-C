#include "pch.h"

using namespace Spire::Doc;

void CreateIfField(intrusive_ptr<Document> document, intrusive_ptr<Paragraph> paragraph)
{
	intrusive_ptr<IfField> ifField = new IfField(document);
	ifField->SetType(FieldType::FieldIf);
	ifField->SetCode(L"IF ");

	paragraph->GetItems()->Add(ifField);
	paragraph->AppendField(L"Count", FieldType::FieldMergeField);
	paragraph->AppendText(L" > ");
	paragraph->AppendText(L"\"100\" ");
	paragraph->AppendText(L"\"Thanks\" ");
	paragraph->AppendText(L"\"The minimum order is 100 units\"");

	intrusive_ptr<ParagraphBase> end = document->CreateParagraphItem(ParagraphItemType::FieldMark);
	intrusive_ptr<FieldMark> fm = Object::Dynamic_cast<FieldMark>(end);
	fm->SetType(FieldMarkType::FieldEnd);
	paragraph->GetItems()->Add(end);
	ifField->SetEnd(Object::Dynamic_cast<FieldMark>(end));
}

int main()
{
	std::wstring outputFile = OUTPUTPATH"/CreateIFField.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();

	//Add a new section.
	intrusive_ptr<Section> section = document->AddSection();

	//Add a new paragraph.
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	// Define a method of creating an IF Field.
	CreateIfField(document, paragraph);

	//Define merged data.
	std::vector<LPCWSTR_S> fieldName = { L"Count" };
	std::vector<LPCWSTR_S> fieldValue = { L"2" };

	//Merge data into the IF Field.
	document->GetMailMerge()->Execute(fieldName, fieldValue);

	//Update all fields in the document.
	document->SetIsUpdateFields(true);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}