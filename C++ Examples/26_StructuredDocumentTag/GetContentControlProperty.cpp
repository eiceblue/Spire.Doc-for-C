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

intrusive_ptr<StructureTagsInternal> GetAllTags(intrusive_ptr<Document> document) {
	intrusive_ptr<StructureTagsInternal> structureTags = new StructureTagsInternal();
	for (int i = 0; i < document->GetSections()->GetCount(); i++) {
		intrusive_ptr<Section> section = document->GetSections()->GetItemInSectionCollection(i);
		for (int j = 0; j < section->GetBody()->GetChildObjects()->GetCount(); j++) {
			intrusive_ptr<DocumentObject> obj = section->GetBody()->GetChildObjects()->GetItem(j);
			if (obj->GetDocumentObjectType() == DocumentObjectType::StructureDocumentTag)
			{
				intrusive_ptr<StructureDocumentTag> tagObj = Object::Dynamic_cast<StructureDocumentTag>(obj);
				structureTags->GetTags().push_back(tagObj);
			}
			else if (obj->GetDocumentObjectType() == DocumentObjectType::Paragraph)
			{
				for (int k = 0; k < (Object::Dynamic_cast<Paragraph>(obj))->GetChildObjects()->GetCount(); k++)
				{
					intrusive_ptr<DocumentObject> pobj = (Object::Dynamic_cast<Paragraph>(obj))->GetChildObjects()->GetItem(k);
					if (pobj->GetDocumentObjectType() == DocumentObjectType::StructureDocumentTagInline)
					{
						structureTags->GetTagInlines().push_back(Object::Dynamic_cast<StructureDocumentTagInline>(pobj));
					}
				}
			}
			else if (obj->GetDocumentObjectType() == DocumentObjectType::Table)
			{
				for (int k = 0; k < (Object::Dynamic_cast<Table>(obj))->GetRows()->GetCount(); k++)
				{
					intrusive_ptr<TableRow> row = (Object::Dynamic_cast<Table>(obj))->GetRows()->GetItemInRowCollection(k);
					for (int l = 0; l < row->GetCells()->GetCount(); l++)
					{
						intrusive_ptr<TableCell> cell = row->GetCells()->GetItemInCellCollection(l);
						for (int m = 0; m < cell->GetChildObjects()->GetCount(); m++)
						{
							intrusive_ptr<DocumentObject> cellChild = cell->GetChildObjects()->GetItem(m);
							if (cellChild->GetDocumentObjectType() == DocumentObjectType::StructureDocumentTag)
							{
								structureTags->GetTags().push_back(Object::Dynamic_cast<StructureDocumentTag>(cellChild));
							}
							else if (cellChild->GetDocumentObjectType() == DocumentObjectType::Paragraph)
							{
								for (int n = 0; n < (Object::Dynamic_cast<Paragraph>(cellChild))->GetChildObjects()->GetCount(); n++)
								{
									intrusive_ptr<DocumentObject> pobj = (Object::Dynamic_cast<Paragraph>(cellChild))->GetChildObjects()->GetItem(n);
									if (pobj->GetDocumentObjectType() == DocumentObjectType::StructureDocumentTagInline)
									{
										structureTags->GetTagInlines().push_back(Object::Dynamic_cast<StructureDocumentTagInline>(pobj));
									}
								}
							}
						}
					}
				}
			}
		}
	}
	//C# TO C++ CONVERTER TODO TASK: A 'delete structureTags' statement was not added since structureTags was used in a 'return' or 'throw' statement.
	return structureTags;
}
wstring GetSDTType(SdtType value)
{
	switch (value)
	{
	case Spire::Doc::SdtType::None:
		return L"None";
		break;
	case Spire::Doc::SdtType::RichText:
		return L"RichText";
		break;
	case Spire::Doc::SdtType::Bibliography:
		return L"Bibliography";
		break;
	case Spire::Doc::SdtType::Citation:
		return L"Citation";
		break;
	case Spire::Doc::SdtType::ComboBox:
		return L"ComboBox";
		break;
	case Spire::Doc::SdtType::DropDownList:
		return L"DropDownList";
		break;
	case Spire::Doc::SdtType::Picture:
		return L"Picture";
		break;
	case Spire::Doc::SdtType::Text:
		return L"Text";
		break;
	case Spire::Doc::SdtType::Equation:
		return L"Equation";
		break;
	case Spire::Doc::SdtType::DatePicker:
		return L"DatePicker";
		break;
	case Spire::Doc::SdtType::BuildingBlockGallery:
		return L"BuildingBlockGallery";
		break;
	case Spire::Doc::SdtType::DocPartObj:
		return L"DocPartObj";
		break;
	case Spire::Doc::SdtType::Group:
		return L"Group";
		break;
	case Spire::Doc::SdtType::CheckBox:
		return L"CheckBox";
		break;
	}
	return L"";
}


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ContentControl.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"GetContentControlProperty.txt";
	

	//Create document and load file from disk
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get all structureTags in the Word document
	intrusive_ptr<StructureTagsInternal> structureTags = GetAllTags(doc);
	//Get all StructureDocumentTagInline objects
	std::vector<intrusive_ptr<StructureDocumentTagInline>> tagInlines = structureTags->GetTagInlines();
	std::wstring stringBuidler;
	stringBuidler.append(L"Alias of contentControl")
		.append(L"\t")
		.append(L"ID          ")
		.append(L"\t")
		.append(L"Tag             ")
		.append(L"\t")
		.append(L"STDType        ")
		.append(L"\n");
	//Get properties of all tagInlines
	for (size_t i = 0; i < tagInlines.size(); i++) {
		intrusive_ptr<StructureDocumentTagInline> tagInline = tagInlines[i];
		std::wstring alias = tagInline->GetSDTProperties()->GetAlias();
		double id = tagInline->GetSDTProperties()->GetId();
		std::wstring tag = tagInline->GetSDTProperties()->GetTag();
		//C# TO C++ CONVERTER TODO TASK: There is no C++ equivalent to 'ToString':
		std::wstring STDType = GetSDTType(tagInline->GetSDTProperties()->GetSDTType());
		stringBuidler.append(alias)
			.append(L",\t")
			.append(to_wstring(id))
			.append(L",\t")
			.append(tag)
			.append(L",\t")
			.append(STDType)
			.append(L"\n");
	}


	//Get all StructureDocumentTag objects
	std::vector<intrusive_ptr<StructureDocumentTag>> tags = structureTags->GetTags();
	//Get properties of all tags
	for (size_t i = 0; i < tags.size(); i++) {
		intrusive_ptr<StructureDocumentTag> tag = tags[i];
		std::wstring alias = tag->GetSDTProperties()->GetAlias();
		double id = tag->GetSDTProperties()->GetId();
		std::wstring tagStr = tag->GetSDTProperties()->GetTag();
		//C# TO C++ CONVERTER TODO TASK: There is no C++ equivalent to 'ToString':
		std::wstring STDType = GetSDTType(tag->GetSDTProperties()->GetSDTType());
		stringBuidler.append(alias)
			.append(L",\t")
			.append(to_wstring(id))
			.append(L",\t")
			.append(tagStr)
			.append(L",\t")
			.append(STDType)
			.append(L"\n");
	}


	//Save the property to a text document and launch it
	wofstream out1;
	out1.open(outputFile.c_str());
	out1.flush();
	out1 << stringBuidler.c_str();
	out1.close();
	
	doc->Close();
}