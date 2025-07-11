#include "pch.h"
using namespace Spire::Doc;

class StructureTagsInternal : public ReferenceCounter
{
private:
	std::vector<intrusive_ptr<StructureDocumentTagInline>> m_tagInlines;
	std::vector<intrusive_ptr<StructureDocumentTag>> m_tags;

public:
	std::vector<intrusive_ptr<StructureDocumentTagInline>>& GetTagInlines()
	{
		if (m_tagInlines.empty())
		{
			m_tagInlines = std::vector<intrusive_ptr<StructureDocumentTagInline>>();
		}
		return m_tagInlines;
	}
	void SetTagInlines(std::vector<intrusive_ptr<StructureDocumentTagInline>> value)
	{
		m_tagInlines = value;
	}
	std::vector<intrusive_ptr<StructureDocumentTag>>& GetTags()
	{
		if (m_tags.empty())
		{
			m_tags = std::vector<intrusive_ptr<StructureDocumentTag>>();
		}
		return m_tags;
	}
	void SetTags(std::vector<intrusive_ptr<StructureDocumentTag>> value)
	{
		m_tags = value;
	}
};

intrusive_ptr<StructureTagsInternal> GetAllTags(intrusive_ptr<Document> document)
{

	//Create StructureTags

	intrusive_ptr<StructureTagsInternal> structureTags = new StructureTagsInternal();

	//Travel document sections
	for (int i = 0; i < document->GetSections()->GetCount(); i++)
	{
		intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++)
		{
			intrusive_ptr<DocumentObject> obj = section->GetBody()->GetChildObjects()->GetItem(j);
			//Travel document paragraphs
			if (obj->GetDocumentObjectType() == DocumentObjectType::Paragraph)
			{
				for (int k = 0; k < (Object::Dynamic_cast <Paragraph> (obj))->GetChildObjects()->GetCount(); k++)
				{
					intrusive_ptr<DocumentObject> pobj = (Object::Dynamic_cast<Paragraph>(obj))->GetChildObjects()->GetItem(k);
					//Get StructureDocumentTagInline
					if (pobj->GetDocumentObjectType() == DocumentObjectType::StructureDocumentTagInline)
					{
						structureTags->GetTagInlines().push_back(Object::Dynamic_cast<StructureDocumentTagInline>(pobj));
					}
				}
			}

		}
	}

	return structureTags;
}

wstring GetStdTypeStr(SdtType value)
{
	wstring ret = L"";
	switch (value)
	{
	case Spire::Doc::SdtType::None:
		ret = L"None";
		break;
	case Spire::Doc::SdtType::RichText:
		ret = L"RichText";
		break;
	case Spire::Doc::SdtType::Bibliography:
		ret = L"Bibliography";
		break;
	case Spire::Doc::SdtType::Citation:
		ret = L"Citation";
		break;
	case Spire::Doc::SdtType::ComboBox:
		ret = L"ComboBox";
		break;
	case Spire::Doc::SdtType::DropDownList:
		ret = L"DropDownList";
		break;
	case Spire::Doc::SdtType::Picture:
		ret = L"Picture";
		break;
	case Spire::Doc::SdtType::Text:
		ret = L"Text";
		break;
	case Spire::Doc::SdtType::Equation:
		ret = L"Equation";
		break;
	case Spire::Doc::SdtType::DatePicker:
		ret = L"DatePicker";
		break;
	case Spire::Doc::SdtType::BuildingBlockGallery:
		ret = L"BuildingBlockGallery";
		break;
	case Spire::Doc::SdtType::DocPartObj:
		ret = L"DocPartObj";
		break;
	case Spire::Doc::SdtType::Group:
		ret = L"Group";
		break;
	case Spire::Doc::SdtType::CheckBox:
		ret = L"CheckBox";
		break;
	default:
		break;
	}
	return ret;
}


int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring inputFile = input_path + L"CheckBoxContentControl.docx";
	wstring outputFile = output_path + L"UpdateCheckBox.docx";

	//Create a document
	intrusive_ptr<Document> document = new Document();

	//Load the document from disk.
	document->LoadFromFile(inputFile.c_str());

	//Call StructureTags
	intrusive_ptr<StructureTagsInternal> structureTags = GetAllTags(document);

	//Create list 
	vector<intrusive_ptr<StructureDocumentTagInline>>& tagInlines = structureTags->GetTagInlines();

	//Get the controls
	for (size_t i = 0; i < tagInlines.size(); i++)
	{
		//Get the type
		wstring type = GetStdTypeStr(tagInlines[i]->GetSDTProperties()->GetSDTType());

		//Update the status
		if (type == L"CheckBox")
		{
			intrusive_ptr<SdtCheckBox> scb = Object::Dynamic_cast<SdtCheckBox>(tagInlines[i]->GetSDTProperties()->GetControlProperties());
			if (scb->GetChecked())
			{
				scb->SetChecked(false);
			}
			else
			{
				scb->SetChecked(true);
			}
		}

	}
	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();


}

