#include "pch.h"

using namespace Spire::Doc;

std::vector<intrusive_ptr<Field>> FindAllHyperlinks(intrusive_ptr<Document> document)
{
	std::vector<intrusive_ptr<Field>> hyperlinks;
	//Iterate through the items in the sections to find all hyperlinks
	int sectionCount = document->GetSections()->GetCount();
	for (int i = 0; i < sectionCount; i++)
	{
		intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
		int secBodyChildCount = section->GetBody()->GetChildObjects()->GetCount();
		for (int j = 0; j < secBodyChildCount; j++)
		{
			intrusive_ptr<DocumentObject> childObj = section->GetBody()->GetChildObjects()->GetItem(j);
			if (childObj->GetDocumentObjectType() == DocumentObjectType::Paragraph)
			{
				int paraChildCount = (Object::Dynamic_cast<Paragraph>(childObj))->GetChildObjects()->GetCount();
				for (int k = 0; k < paraChildCount; k++)
				{
					intrusive_ptr<DocumentObject> paraObj = (Object::Dynamic_cast<Paragraph>(childObj))->GetChildObjects()->GetItem(k);
					if (paraObj->GetDocumentObjectType() == DocumentObjectType::Field)
					{
						intrusive_ptr<Field> field = Object::Dynamic_cast<Field>(paraObj);
						if (field->GetType() == FieldType::FieldHyperlink)
						{
							hyperlinks.push_back(field);
						}
					}
				}
			}
		}
	}
	return hyperlinks;
}

void FlattenHyperlinks(intrusive_ptr<Field> field)
{
	int ownerParaIndex = field->GetOwnerParagraph()->GetOwnerTextBody()->GetChildObjects()->IndexOf(field->GetOwnerParagraph());
	int fieldIndex = field->GetOwnerParagraph()->GetChildObjects()->IndexOf(field);
	intrusive_ptr<Paragraph> sepOwnerPara = field->GetSeparator()->GetOwnerParagraph();
	int sepOwnerParaIndex = field->GetSeparator()->GetOwnerParagraph()->GetOwnerTextBody()->GetChildObjects()->IndexOf(field->GetSeparator()->GetOwnerParagraph());
	int sepIndex = field->GetSeparator()->GetOwnerParagraph()->GetChildObjects()->IndexOf(field->GetSeparator());
	int endIndex = field->GetEnd()->GetOwnerParagraph()->GetChildObjects()->IndexOf(field->GetEnd());
	int endOwnerParaIndex = field->GetEnd()->GetOwnerParagraph()->GetOwnerTextBody()->GetChildObjects()->IndexOf(field->GetEnd()->GetOwnerParagraph());

	FormatFieldResultText(field->GetSeparator()->GetOwnerParagraph()->GetOwnerTextBody(), sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex);

	field->GetEnd()->GetOwnerParagraph()->GetChildObjects()->RemoveAt(endIndex);

	for (int i = sepOwnerParaIndex; i >= ownerParaIndex; i--)
	{
		if (i == sepOwnerParaIndex && i == ownerParaIndex)
		{
			for (int j = sepIndex; j >= fieldIndex; j--)
			{
				field->GetOwnerParagraph()->GetChildObjects()->RemoveAt(j);

			}
		}
		else if (i == ownerParaIndex)
		{
			for (int j = field->GetOwnerParagraph()->GetChildObjects()->GetCount() - 1; j >= fieldIndex; j--)
			{
				field->GetOwnerParagraph()->GetChildObjects()->RemoveAt(j);
			}

		}
		else if (i == sepOwnerParaIndex)
		{
			for (int j = sepIndex; j >= 0; j--)
			{
				sepOwnerPara->GetChildObjects()->RemoveAt(j);
			}
		}
		else
		{
			field->GetOwnerParagraph()->GetOwnerTextBody()->GetChildObjects()->RemoveAt(i);
		}
	}
}

void FormatFieldResultText(intrusive_ptr<Body> ownerBody, int sepOwnerParaIndex, int endOwnerParaIndex, int sepIndex, int endIndex)
{
	for (int i = sepOwnerParaIndex; i <= endOwnerParaIndex; i++)
	{
		intrusive_ptr<Paragraph> para = Object::Dynamic_cast<Paragraph>(ownerBody->GetChildObjects()->GetItem(i));
		if (i == sepOwnerParaIndex && i == endOwnerParaIndex)
		{
			for (int j = sepIndex + 1; j < endIndex; j++)
			{
				FormatText(Object::Dynamic_cast<TextRange>(para->GetChildObjects()->GetItem(j)));
			}

		}
		else if (i == sepOwnerParaIndex)
		{
			for (int j = sepIndex + 1; j < para->GetChildObjects()->GetCount(); j++)
			{
				FormatText(Object::Dynamic_cast<TextRange>(para->GetChildObjects()->GetItem(j)));
			}
		}
		else if (i == endOwnerParaIndex)
		{
			for (int j = 0; j < endIndex; j++)
			{
				FormatText(Object::Dynamic_cast<TextRange>(para->GetChildObjects()->GetItem(j)));
			}
		}
		else
		{
			for (int j = 0; j < para->GetChildObjects()->GetCount(); j++)
			{
				FormatText(Object::Dynamic_cast<TextRange>(para->GetChildObjects()->GetItem(j)));
			}
		}
	}
}

void FormatText(intrusive_ptr<TextRange> tr)
{
	//Set the text color to black
	tr->GetCharacterFormat()->SetTextColor(Color::GetBlack());
	//Set the text underline style to none
	tr->GetCharacterFormat()->SetUnderlineStyle(UnderlineStyle::None);
}

int main()
{
	std::wstring outputFile = OUTPUTPATH"/RemoveHyperlinks.docx";
	std::wstring inputFile = DATAPATH"/Hyperlinks.docx";

	//Load Document
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get all hyperlinks
	std::vector<intrusive_ptr<Field>> hyperlinks = FindAllHyperlinks(doc);

	//Flatten all hyperlinks
	for (int i = hyperlinks.size() - 1; i >= 0; i--)
	{
		FlattenHyperlinks(hyperlinks[i]);
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
}