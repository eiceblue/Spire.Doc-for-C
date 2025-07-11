#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ConvertDocToByte.docx";

	intrusive_ptr<Document> doc = new Document();
	// Load the document from disk.
	doc->LoadFromFile(inputFile.c_str());

	// Create a new stream.
	intrusive_ptr<Stream> outStream = new Stream();

	// Save the document to stream.
	doc->SaveToStream(outStream, FileFormat::Docx);

	// Convert the document to bytes.
	std::vector<byte> docBytes = outStream->ToArray();

	// The bytes are now ready to be stored/transmitted.

	// Now reverse the steps to load the bytes back into a document object.
	intrusive_ptr<Stream> inStream = new Stream(docBytes.data(), docBytes.size());


	// Load the stream into a new document object.
	intrusive_ptr<Document>  newDoc = new Document(inStream);
	//save doc file.
	intrusive_ptr<Stream> ms = new Stream();

	newDoc->SaveToStream(ms, FileFormat::Docx);
	std::ofstream outFile(outputFile, std::ios::out | std::ofstream::binary);
	std::vector<byte> data = ms->ToArray();
	outFile.write((char*)(&data[0]), data.size() * sizeof(byte));
	outFile.close();

	newDoc->Close();
}