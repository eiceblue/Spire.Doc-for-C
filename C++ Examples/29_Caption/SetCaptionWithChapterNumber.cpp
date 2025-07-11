#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SetCaptionWithChapterNumber.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetCaptionWithChapterNumber.docx";

	//Create word document
	intrusive_ptr<Document> document = new Document();
	//Load file
	document->LoadFromFile(inputFile.c_str());
	//Get the first section
	intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(0);
	//Label name
	std::wstring name = L"Caption ";
	for (int i = 0; i < section->GetBody()->GetParagraphs()->GetCount(); i++)
	{
		for (int j = 0; j < section->GetBody()->GetParagraphs()->GetItemInParagraphCollection(i)->GetChildObjects()->GetCount(); j++)
		{
			if (Object::CheckType<DocPicture>(section->GetBody()->GetParagraphs()->GetItemInParagraphCollection(i)->GetChildObjects()->GetItem(j)))
			{
				intrusive_ptr<DocPicture> pic1 = boost::dynamic_pointer_cast<DocPicture>(section->GetBody()->GetParagraphs()->GetItemInParagraphCollection(i)->GetChildObjects()->GetItem(j));
				intrusive_ptr<Body> body = Object::Dynamic_cast<Body>(pic1->GetOwnerParagraph()->GetOwner());
				if (body != nullptr)
				{
					int imageIndex = body->GetChildObjects()->IndexOf(pic1->GetOwnerParagraph());
					//Create a new paragraph
					intrusive_ptr<Paragraph> para = new Paragraph(document);
					//Set label
					para->AppendText(name.c_str());

					//Add caption
					intrusive_ptr<Field> field1 = para->AppendField(L"test", FieldType::FieldStyleRef);
					//Chapter number
					field1->SetCode(L" STYLEREF 1 \\s ");
					//Chapter delimiter
					para->AppendText(L" - ");

					//Add picture sequence number
					intrusive_ptr<SequenceField> field2 = Object::Dynamic_cast<SequenceField>(para->AppendField(name.c_str(), FieldType::FieldSequence));
					field2->SetCaptionName(name.c_str());
					field2->SetNumberFormat(CaptionNumberingFormat::Number);
					body->GetParagraphs()->Insert(imageIndex + 1, para);
				}
			}
		}
	}
	//Set update fields
	document->SetIsUpdateFields(true);
	//Save the result file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
}
