import axios from 'axios';

export const getGraphData = async (url: string, accesstoken: string) => {
    const response = await axios({
        url: url,
        method: 'get',
        headers: {'Authorization': `Bearer ${accesstoken}`}
      });
    return response;
};

export const createMailFolder = async (accesstoken: string) => {
  try {
    const response = await axios({
      url: 'https://graph.microsoft.com/v1.0/me/mailFolders',
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accesstoken}`,
        'Content-Type': 'application/json'
      },
      data: {
        "displayName": "test",
        "isHidden": false
      }
    });
    console.log('Mail folder created successfully:', response.data);
    return response.data;
  } catch (error) {
    console.error('Error creating mail folder:',  error);
    throw error;
  }
};
