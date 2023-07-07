import { sp } from "@pnp/sp/presets/all";

// Configure SharePoint context
sp.setup({
  sp: {
    baseUrl: "https://rtuacin.sharepoint.com/sites/shivang"
  }
});

// Upload file to SharePoint document library
export const uploadFileToLibrary = async (fileName: string, file: File) => {
    try {
      const documentLibraryRelativeUrl = "/sites/shivang/Shared Documents"; // Relative URL of the document library
      const folder = sp.web.getFolderByServerRelativeUrl(documentLibraryRelativeUrl);
      
      // Upload the file
      await folder.files.add(fileName, file, true);
  
      console.log("File uploaded successfully.");
    } catch (error) {
      console.log(`Error uploading file: ${error}`);
    }
  };
  