#include "pch.h"
#include <deque>
using namespace Spire::Doc;

int main()
{
	std::wstring outputFile = OUTPUTPATH"/ExtractImage/";
	std::wstring inputFile = DATAPATH"/Template.docx";

	//open document
	intrusive_ptr<Document> document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//document elements, each of them has child elements
	std::deque<intrusive_ptr<ICompositeObject>> nodes;
	nodes.push_back(document);

	//embedded images list.
	std::vector<std::vector<byte>> images;
	//traverse
	while (nodes.size() > 0)
	{
		intrusive_ptr<ICompositeObject> node = nodes.front();
		nodes.pop_front();
		for (int i = 0; i < node->GetChildObjects()->GetCount(); i++)
		{
			intrusive_ptr<IDocumentObject> child = node->GetChildObjects()->GetItem(i);
			if (child->GetDocumentObjectType() == DocumentObjectType::Picture)
			{
				intrusive_ptr<DocPicture> picture = Object::Dynamic_cast<DocPicture>(child);
				std::vector<byte> imageByte = picture->GetImageBytes();
				images.push_back(imageByte);
			}
			else if (Object::CheckType<ICompositeObject>(child))
			{
				nodes.push_back(boost::dynamic_pointer_cast<ICompositeObject>(child));
			}
		}
	}
	//save images
	for (size_t i = 0; i < images.size(); i++)
	{
		std::wstring fileName = L"Image-" + to_wstring(i) + L".png";
		std::wstring tempImagePath = outputFile + fileName;
		std::ofstream outFile(tempImagePath, std::ios::binary);
		if (outFile.is_open())
		{
			outFile.write(reinterpret_cast<const char*>(images[i].data()), images[i].size());
			outFile.close();
		}
	}
	document->Close();
}