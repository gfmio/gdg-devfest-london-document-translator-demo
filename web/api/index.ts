import Axios from "axios";

const API_BASE_URL = process.env.API_BASE_URL;

export interface Language {
  code: string;
  name: string;
}

const getSupportedLanguages = async (): Promise<Language[]> => {
  const response = await Axios.get(`${API_BASE_URL}/supported-languages`);
  return response.data;
}

const translateDocument = async (document: File, to: string, from?: string | null): Promise<ArrayBuffer> => {
  const formData = new FormData();
  formData.append("document", document);
  if (from) {
    formData.append("from", from);
  }
  formData.append("to", to);

  const response = await Axios.post(`${API_BASE_URL}/translate-document`, formData, {
    headers: {
      'Content-Type': 'multipart/form-data'
    },
    responseType: "arraybuffer"
  });

  return response.data;
}

export default {
  getSupportedLanguages,
  translateDocument,
}
