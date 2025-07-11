#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"OLEs.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile_pdf = output_path + L"ExtractOLE.pdf";
	wstring outputFile_xls = output_path + L"ExtractOLE.xls";
	wstring outputFile_pptx = output_path + L"ExtractOLE.pptx";

	//Create document and load file from disk
	intrusive_ptr<Document> doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Traverse through all sections of the word document    
	for (int s = 0; s < doc->GetSections()->GetCount(); s++)
	{
		intrusive_ptr<Section> sec = doc->GetSections()->GetItemInSectionCollection(s);
		//Traverse through all Child Objects in the GetBody() of each section
		for (int i = 0; i < sec->GetBody()->GetChildObjects()->GetCount(); i++)
		{
			intrusive_ptr<DocumentObject> obj = sec->GetBody()->GetChildObjects()->GetItem(i);
			//find the paragraph
			if (Object::CheckType<Paragraph>(obj))
			{
				intrusive_ptr<Paragraph> par = boost::dynamic_pointer_cast<Paragraph>(obj);
				for (int j = 0; j < par->GetChildObjects()->GetCount(); j++)
				{
					intrusive_ptr<DocumentObject> o = par->GetChildObjects()->GetItem(j);
					//check whether the object is OLE
					if (o->GetDocumentObjectType() == DocumentObjectType::OleObject)
					{
						intrusive_ptr<DocOleObject> Ole = Object::Dynamic_cast<DocOleObject>(o);
						std::wstring s = Ole->GetObjectType();

						//check whether the object type is "Acrobat.Document.11"
						if (wcscmp(s.c_str(), L"AcroExch.Document.DC") == 0)
						{
							//write the data of OLE into file										
							std::ofstream pdf_file(outputFile_pdf, std::ios::out | std::ofstream::binary);
							std::vector<byte> native_data = Ole->GetNativeData();
							pdf_file.write((char*)(&native_data[0]), native_data.size() * sizeof(byte));
							pdf_file.close();
							
						}

						//check whether the object type is "Excel.Sheet.8"
						else if (wcscmp(s.c_str(), L"Excel.Sheet.8") == 0)
						{
							std::ofstream xls_file(outputFile_xls, std::ios::out | std::ofstream::binary);
							std::vector<byte> native_data = Ole->GetNativeData();
							xls_file.write((char*)(&native_data[0]), native_data.size() * sizeof(byte));
							xls_file.close();
																	
						}

						//check whether the object type is "PowerPoint.Show.12"
						else if (wcscmp(s.c_str(), L"PowerPoint.Show.12") == 0)
						{
							std::ofstream pptx_file(outputFile_pptx, std::ios::out | std::ofstream::binary);
							std::vector<byte> native_data = Ole->GetNativeData();
							pptx_file.write((char*)(&native_data[0]), native_data.size() * sizeof(byte));
							pptx_file.close();
							
						}
					}
				}
			}
		}
	}
	doc->Close();
}