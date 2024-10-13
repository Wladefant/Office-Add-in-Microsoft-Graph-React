const express = require('express');
const cors = require('cors');
const app = express();
const port = process.env.PORT || 3001;

// Enable CORS
app.use(cors());
app.use(express.json()); // Add this line to parse JSON request bodies

const CosmosClient = require('@azure/cosmos').CosmosClient; // Import CosmosClient
const { Configuration, OpenAIApi } = require('openai'); // Import OpenAI


const config = require('./config')

const endpoint = config.endpoint
const key = config.key

const databaseId = config.database.id
const containerId = config.container.id
const emailsContainerId = config.container.emails.id; // Use the correct container ID
const partitionKey = { kind: 'Hash', paths: ['/userId'] };

const options = {
      endpoint: endpoint,
      key: key,
      userAgentSuffix: 'CosmosDBJavascriptQuickstart'
    };

const client = new CosmosClient(options)

// Initialize OpenAI
const openaiConfiguration = new Configuration({
  apiKey: '', // Replace with your OpenAI API key
});
const openai = new OpenAIApi(openaiConfiguration);



// Function to extract location using OpenAI
async function extractLocation(emailContent) {
  try {
    const prompt = `Bestimme den Ort aus dem folgenden E-Mail-Inhalt und gib als Output nur die Adresse wieder. Falls nicht gefunden, schreibe "nicht gefunden". E-Mail-Inhalt: ${emailContent}`;

    const response = await openai.createChatCompletion({
      model: 'gpt-4o', // Use 'gpt-4' if available and needed
      messages: [
        {
          role: 'system',
          content: 'Du bist ein hilfreicher Assistent, der den Ort aus E-Mail-Inhalten extrahiert.',
        },
        {
          role: 'user',
          content: prompt,
        },
      ],
      max_tokens: 50,
      n: 1,
      temperature: 0.7,
    });

    if (response.data.choices && response.data.choices.length > 0) {
      return response.data.choices[0].message.content.trim();
    } else {
      return 'nicht gefunden';
    }
  } catch (error) {
    console.error('Error extracting location:', error);
    return 'nicht gefunden';
  }
}
/**
 * Create the database if it does not exist
 */
async function createDatabase() {
  const { database } = await client.databases.createIfNotExists({
    id: databaseId
  })
  console.log(`Created database:\n${database.id}\n`)
}

/**
 * Read the database definition
 */
async function readDatabase() {
  const { resource: databaseDefinition } = await client
    .database(databaseId)
    .read()
  console.log(`Reading database:\n${databaseDefinition.id}\n`)
}

/**
 * Create the container if it does not exist
 */
async function createContainer() {
  const { container } = await client
    .database(databaseId)
    .containers.createIfNotExists(
      { id: containerId, partitionKey }
    )
  console.log(`Created container:\n${config.container.id}\n`)
}

/**
 * Read the container definition
 */
async function readContainer() {
  const { resource: containerDefinition } = await client
    .database(databaseId)
    .container(containerId)
    .read()
  console.log(`Reading container:\n${containerDefinition.id}\n`)
}

/**
 * Scale a container
 * You can scale the throughput (RU/s) of your container up and down to meet the needs of the workload. Learn more: https://aka.ms/cosmos-request-units
 */
async function scaleContainer() {
  const { resource: containerDefinition } = await client
    .database(databaseId)
    .container(containerId)
    .read();
  
  try
  {
      const {resources: offers} = await client.offers.readAll().fetchAll();
  
      const newRups = 500;
      for (var offer of offers) {
        if (containerDefinition._rid !== offer.offerResourceId)
        {
            continue;
        }
        offer.content.offerThroughput = newRups;
        const offerToReplace = client.offer(offer.id);
        await offerToReplace.replace(offer);
        console.log(`Updated offer to ${newRups} RU/s\n`);
        break;
      }
  }
  catch(err)
  {
      if (err.code == 400)
      {
          console.log(`Cannot read container throuthput.\n`);
          console.log(err.body.message);
      }
      else 
      {
          throw err;
      }
  }
}

/**
 * Create family item if it does not exist
 */
async function createFamilyItem(itemBody) {
  const { item } = await client
    .database(databaseId)
    .container(containerId)
    .items.upsert(itemBody)
  console.log(`Created family item with id:\n${itemBody.id}\n`)
}

/**
 * Query the container using SQL
 */
async function queryContainer() {
  console.log(`Querying container:\n${config.container.id}`)

  // query to return all children in a family
  // Including the partition key value of country in the WHERE filter results in a more efficient query
  const querySpec = {
    query: 'SELECT VALUE r.children FROM root r WHERE r.partitionKey = @country',
    parameters: [
      {
        name: '@country',
        value: 'USA'
      }
    ]
  }

  const { resources: results } = await client
    .database(databaseId)
    .container(containerId)
    .items.query(querySpec)
    .fetchAll()
  for (var queryResult of results) {
    let resultString = JSON.stringify(queryResult)
    console.log(`\tQuery returned ${resultString}\n`)
  }
}

/**
 * Replace the item by ID.
 */
async function replaceFamilyItem(itemBody) {
  console.log(`Replacing item:\n${itemBody.id}\n`)
  // Change property 'grade'
  itemBody.children[0].grade = 6
  const { item } = await client
    .database(databaseId)
    .container(containerId)
    .item(itemBody.id, itemBody.partitionKey)
    .replace(itemBody)
}

/**
 * Delete the item by ID.
 */
async function deleteFamilyItem(itemBody) {
  await client
    .database(databaseId)
    .container(containerId)
    .item(itemBody.id, itemBody.partitionKey)
    .delete(itemBody)
  console.log(`Deleted item:\n${itemBody.id}\n`)
}

/**
 * Cleanup the database and collection on completion
 */
async function cleanup() {
  await client.database(databaseId).delete()
}

/**
 * Exit the app with a prompt
 * @param {string} message - The message to display
 */
function exit(message) {
  console.log(message)
  console.log('Press any key to exit')
  process.stdin.setRawMode(true)
  process.stdin.resume()
  process.stdin.on('data', process.exit.bind(process, 0))
}

// Define the endpoint that triggers createDatabase
app.get('/createFamilyItem', async (req, res) => {
  try {
    await createFamilyItem(config.items.Andersen);
    res.status(200).send('Family added successfully.');
  } catch (error) {
    console.error(error);
    res.status(500).send('Error adding Family.');
  }
});

// Define the endpoint that triggers deleteFamilyItem
app.get('/deleteFamilyItem', async (req, res) => {
  try {
    await deleteFamilyItem(config.items.Andersen);
    res.status(200).send('Family deleted successfully.');
  } catch (error) {
    console.error(error);
    res.status(500).send('Error deleting Family.');
  }
});

// Define the endpoint that triggers queryContainer
app.get('/queryContainer', async (req, res) => {
  try {
    await queryContainer();
    res.status(200).send('Query executed successfully.');
  } catch (error) {
    console.error(error);
    res.status(500).send('Error executing query.');
  }
});

// Define the endpoint that checks if a user exists
app.get('/checkUser', async (req, res) => {
  const email = req.query.email;
  const querySpec = {
    query: 'SELECT * FROM c WHERE c.email = @email',
    parameters: [{ name: '@email', value: email }],
  };
  const { resources: users } = await client.database(databaseId).container('Users').items.query(querySpec).fetchAll();
  res.status(200).json({ exists: users.length > 0 });
});

// Define the endpoint that creates a new user
app.post('/createUser', async (req, res) => {
  const { email } = req.body;
  const newUser = { email, id: email };
  await client.database(databaseId).container('Users').items.create(newUser);
  res.status(201).send('User created successfully.');
});

// Define the endpoint that checks if an email exists in CosmosDB based on outlookEmailId
app.get('/checkEmail', async (req, res) => {
  const outlookEmailId = req.query.outlookEmailId;
  const querySpec = {
    query: 'SELECT * FROM c WHERE c.outlookEmailId = @outlookEmailId',
    parameters: [{ name: '@outlookEmailId', value: outlookEmailId }],
  };
  const { resources: emails } = await client.database(databaseId).container('Emails').items.query(querySpec).fetchAll();
  res.status(200).json({ exists: emails.length > 0 });
});

// Define the endpoint that uploads an email to CosmosDB
app.post('/uploadEmail', async (req, res) => {
  const emailData = req.body; 
  try {
    // Check if 'location' already exists in the email data
    if (!emailData.location || emailData.location === '') {
      // Extract 'location' from 'emailBody' using OpenAI
      const location = await extractLocation(emailData.emailBody);

      // Add 'location' to the email data
      emailData.location = location;
    }

    // Save the email data to the 'Emails' container
    await client
      .database(databaseId)
      .container(emailsContainerId)
      .items.create(emailData);

    res.status(201).send('Email uploaded successfully with location.');
  } catch (error) {
    console.error('Error uploading email:', error);
    res.status(500).send('Error uploading email.');
  }
});

// Start the server
app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});
