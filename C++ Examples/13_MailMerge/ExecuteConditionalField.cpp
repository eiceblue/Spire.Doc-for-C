#include "pch.h"

using namespace Spire::Doc;

void CreateIFField1(intrusive_ptr<Document> document, intrusive_ptr<Paragraph> paragraph)
{
	intrusive_ptr<IfField> ifField = new IfField(document);
	ifField->SetType(FieldType::FieldIf);
	ifField->SetCode(L"IF ");
	paragraph->GetItems()->Add(ifField);

	paragraph->AppendField(L"Count", FieldType::FieldMergeField);
	paragraph->AppendText(L" > ");
	paragraph->AppendText(L"\"1\" ");
	paragraph->AppendText(L"\"Greater than one\" ");
	paragraph->AppendText(L"\"Less than one\"");

	intrusive_ptr<IParagraphBase> end = document->CreateParagraphItem(ParagraphItemType::FieldMark);
	(Object::Dynamic_cast<FieldMark>(end))->SetType(FieldMarkType::FieldEnd);
	paragraph->GetItems()->Add(end);

	ifField->SetEnd(Object::Dynamic_cast<FieldMark>(end));
}


void CreateIFField2(intrusive_ptr<Document> document, intrusive_ptr<Paragraph> paragraph)
{
	intrusive_ptr<IfField> ifField = new IfField(document);
	ifField->SetType(FieldType::FieldIf);
	ifField->SetCode(L"IF ");
	paragraph->GetItems()->Add(ifField);

	paragraph->AppendField(L"Age", FieldType::FieldMergeField);
	paragraph->AppendText(L" > ");
	paragraph->AppendText(L"\"50\" ");
	paragraph->AppendText(L"\"The old man\" ");
	paragraph->AppendText(L"\"The young man\"");

	intrusive_ptr<IParagraphBase> end = document->CreateParagraphItem(ParagraphItemType::FieldMark);
	(Object::Dynamic_cast<FieldMark>(end))->SetType(FieldMarkType::FieldEnd);
	paragraph->GetItems()->Add(end);

	ifField->SetEnd(Object::Dynamic_cast<FieldMark>(end));
}

int main() 
{
	std::wstring outputFile = OUTPUTPATH"/ExecuteConditionalField.docx";

	intrusive_ptr<Document> doc = new Document();
	//Add a new section 
	intrusive_ptr<Section> section = doc->AddSection();
	//Add a new paragraph for a section 
	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	CreateIFField1(doc, paragraph);
	paragraph = section->AddParagraph();
	CreateIFField2(doc, paragraph);

	std::vector<LPCWSTR_S> fieldName = { L"Count", L"Age" };
	std::vector<LPCWSTR_S> fieldValue = { L"2", L"30" };

	doc->GetMailMerge()->Execute(fieldName, fieldValue);
	doc->SetIsUpdateFields(true);

	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}



