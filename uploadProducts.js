const axios = require('axios');
const xlsx = require('xlsx');
const slugify = require('slugify');

const STRAPI_URL = 'http://localhost:1337'; // change this if hosted
const STRAPI_TOKEN = '6a92e8f0713555c28f03b59af628d65d9b1e42f380a696eac95c1567fd81945e6521ce591acf3a216730ca440fd81b3d47475bc63ca639ad860238b8b1d76153402d3db5d0ea88968d500ac803edf4897aa0f0cb6005a8e387cd2a38bf1dbdeb32f753381aa97e7bc2e827257ad2ad49f4c1aa6addaa15389a80e4b130e6f2b0'; // paste your admin token here

const api = axios.create({
  baseURL: STRAPI_URL,
  headers: {
    Authorization: `Bearer ${STRAPI_TOKEN}`
  }
});

async function getOrCreateCategory(name) {
  if (!name) return null;
  const res = await api.get('/api/categories', {
    params: { filters: { name: { $eq: name } } }
  });

  if (res.data.data.length > 0) {
    return res.data.data[0].id;
  }

  // Generate a slug for the category
  const slug = slugify(name, { lower: true, strict: true });

  const created = await api.post('/api/categories', {
    data: { name, slug }
  });

  return created.data.data.id;
}
async function generateUniqueSlug(baseSlug) {
  let slug = baseSlug;
  let count = 1;

  while (true) {
    const res = await api.get('/api/products', {
      params: { filters: { slug: { $eq: slug } } }
    });

    if (res.data.data.length === 0) break;

    slug = `${baseSlug}-${count}`;
    count++;
  }

  return slug;
}

async function uploadProduct({ Category, Code, Title, Size }) {
  const name = Title?.trim();
  const code = Code?.trim();
  const size = Size?.trim() || '';
  const categoryName = Category?.trim();

  if (!name || !code) {
    console.warn(`⚠️ Skipping product with missing name or code:`, { Title, Code });
    return;
  }

  const baseSlug = slugify(name, { lower: true, strict: true });
  if (!baseSlug) {
    console.warn(`⚠️ Skipping product with invalid slug:`, { Title });
    return;
  }
  const slug = await generateUniqueSlug(baseSlug);

  if (!slug) {
    console.warn(`⚠️ Skipping product with null slug:`, { Title, baseSlug, slug });
    return;
  }

  const categoryId = await getOrCreateCategory(categoryName);

  const productData = {
    name,
    code,
    size,
    slug,
    category: 1
  };

  // Debug: log the product data before uploading
  console.log('Uploading product:', productData);

  try {
    const res = await api.post('/api/products', { data: productData });
    console.log(`✅ Uploaded: ${name}`);
  } catch (err) {
    console.error(`❌ Failed to upload ${name}:`, err.response?.data || err.message);
  }
}

function importFromXlsx(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rawData = xlsx.utils.sheet_to_json(sheet, { defval: '' });

  // Remove rows where Code or Title is missing or is the header
  const products = rawData.filter(row =>
    row.Code && row.Title && row.Code !== 'Code' && row.Title !== 'Title'
  ).map(row => ({
    Category: row.Category,
    Code: row.Code,
    Title: row.Title,
    Size: row.Size
  }));

  return products;
}

async function main() {
  const products = importFromXlsx('products.xlsx');
  console.log(`📦 Found ${products.length} products. Uploading...`);
  for (const product of products) {
    await uploadProduct(product);
  }
  console.log('✅ Upload complete.');
}

main();
