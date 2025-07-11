#include "pch.h"

using namespace Spire::Doc;

int main()
{
	std::wstring inputFile = DATAPATH"/SampleB_2.docx";
	std::wstring outputFile = OUTPUTPATH"/CreateNestedField.docx";

	//Create Word document.
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Get the first section
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);

	intrusive_ptr<Paragraph> paragraph = section->AddParagraph();

	//Create an IF field
	intrusive_ptr<IfField> ifField = new IfField(document);
	ifField->SetType(FieldType::FieldIf);
	ifField->SetCode(L"IF ");
	paragraph->GetItems()->Add(ifField);

	//Create the embedded IF field
	intrusive_ptr<IfField> ifField2 = new IfField(document);
	ifField2->SetType(FieldType::FieldIf);
	ifField2->SetCode(L"IF ");
	paragraph->GetChildObjects()->Add(ifField2);
	paragraph->GetItems()->Add(ifField2);
	paragraph->AppendText(L"\"200\" < \"50\"   \"200\" \"50\" ");
	intrusive_ptr<IParagraphBase> embeddedEnd = document->CreateParagraphItem(ParagraphItemType::FieldMark);
	(Object::Dynamic_cast<FieldMark>(embeddedEnd))->SetType(FieldMarkType::FieldEnd);
	paragraph->GetItems()->Add(embeddedEnd);
	ifField2->SetEnd(Object::Dynamic_cast<FieldMark>(embeddedEnd));

	paragraph->AppendText(L" > ");
	paragraph->AppendText(L"\"100\" ");
	paragraph->AppendText(L"\"Thanks\" ");
	paragraph->AppendText(L"\"The minimum order is 100 units\"");
	intrusive_ptr<IParagraphBase> end = document->CreateParagraphItem(ParagraphItemType::FieldMark);
	(Object::Dynamic_cast<FieldMark>(end))->SetType(FieldMarkType::FieldEnd);
	paragraph->GetItems()->Add(end);
	ifField->SetEnd(Object::Dynamic_cast<FieldMark>(end));

	//Update all fields in the document.
	document->SetIsUpdateFields(true);

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
}
