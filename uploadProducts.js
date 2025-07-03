const XLSX = require('xlsx');
const axios = require('axios');

// Configuration
const STRAPI_URL = 'http://localhost:1337'; // Change this to your Strapi URL
const STRAPI_TOKEN = '61203c16fb1450a1b21c285e797de58ba13106cc1cbed92d2e6bc67c4b3d17ff18659d5ec9094e3fc944cf7bdce5d1faddaaa222f3f75de72d93643f56564b9359a7ff02cb0a0ea627eba4a2b7030948198b1e53751d1221bab4b5fd281b098ba9e4dec24762c25109b60c4bc9ed98a1f856d69fd89d051cb8b81152785cd6c9'; // Replace with your actual API token
const EXCEL_FILE_PATH = './products.xlsx'; // Replace with your Excel file path


// Helper function to create slug from name
function createSlug(name) {
  return name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

// Helper function to check if slug exists
async function checkSlugExists(slug) {
  try {
    const response = await makeRequest('GET', `/products?filters[slug][$eq]=${encodeURIComponent(slug)}`);
    return response.data.length > 0;
  } catch (error) {
    console.error('Error checking slug:', error);
    return false;
  }
}

// Helper function to generate unique slug
async function generateUniqueSlug(name) {
  let baseSlug = createSlug(name);
  let uniqueSlug = baseSlug;
  let counter = 1;
  
  // Keep checking until we find a unique slug
  while (await checkSlugExists(uniqueSlug)) {
    uniqueSlug = `${baseSlug}-${counter}`;
    counter++;
  }
  
  return uniqueSlug;
}

// Helper function to make API requests
async function makeRequest(method, endpoint, data = null) {
  const config = {
    method,
    url: `${STRAPI_URL}/api${endpoint}`,
    headers: {
      'Authorization': `Bearer ${STRAPI_TOKEN}`,
      'Content-Type': 'application/json',
    },
  };

  if (data) {
    config.data = data;
  }

  try {
    const response = await axios(config);
    return response.data;
  } catch (error) {
    console.error(`Error making ${method} request to ${endpoint}:`, error.response?.data || error.message);
    throw error;
  }
}

// Function to get subcategory by name
async function getSubCategoryByName(name) {
  try {
    const response = await makeRequest('GET', `/sub-categories?filters[name][$eq]=${encodeURIComponent(name)}`);
    return response.data.length > 0 ? response.data[0] : null;
  } catch (error) {
    console.error('Error fetching subcategory:', error);
    return null;
  }
}

// Function to check if product exists by code
async function getProductByCode(code) {
  try {
    const response = await makeRequest('GET', `/products?filters[code][$eq]=${encodeURIComponent(code)}`);
    return response.data.length > 0 ? response.data[0] : null;
  } catch (error) {
    console.error('Error checking product:', error);
    return null;
  }
}

// Function to create product
async function createProduct(productData, subCategoryDocumentId) {
  // Generate unique slug
  const uniqueSlug = await generateUniqueSlug(productData.name);
  
  const productPayload = {
    data: {
      name: productData.name,
      code: productData.code,
      size: productData.size || '',
      slug: uniqueSlug,
      sub_category: subCategoryDocumentId, // Link to subcategory
      publishedAt: new Date().toISOString(), // Auto-publish
    }
  };

  try {
    const response = await makeRequest('POST', '/products', productPayload);
    console.log(`✅ Created product: ${productData.name} (${productData.code}) with slug: ${uniqueSlug}`);
    return response.data;
  } catch (error) {
    console.error(`❌ Failed to create product: ${productData.name}`, error);
    throw error;
  }
}

// Function to update existing product
async function updateProduct(productDocumentId, productData, subCategoryDocumentId) {
  // Generate unique slug (in case the name changed)
  const uniqueSlug = await generateUniqueSlug(productData.name);
  
  const productPayload = {
    data: {
      name: productData.name,
      code: productData.code,
      size: productData.size || '',
      slug: uniqueSlug,
      sub_category: subCategoryDocumentId, // Link to subcategory
    }
  };

  try {
    const response = await makeRequest('PUT', `/products/${productDocumentId}`, productPayload);
    console.log(`✅ Updated product: ${productData.name} (${productData.code}) with slug: ${uniqueSlug}`);
    return response.data;
  } catch (error) {
    console.error(`❌ Failed to update product: ${productData.name}`, error);
    throw error;
  }
}

// Main function to process Excel file
async function processProductsFile() {
  try {
    console.log('📖 Reading products Excel file...');
    
    // Read the Excel file
    const workbook = XLSX.readFile(EXCEL_FILE_PATH);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert to JSON with header row
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    console.log(`📊 Found ${data.length} rows in Excel file`);
    
    // Get headers from first row
    const headers = data[0];
    console.log('📋 Headers found:', headers);
    
    // Process data rows (skip header)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Map row data to columns
      const subCategoryName = row[0]?.toString().trim(); // Sub Category
      const productCode = row[1]?.toString().trim();     // Code
      const productTitle = row[2]?.toString().trim();    // Title
      const productSize = row[3]?.toString().trim();     // Size
      
      if (!subCategoryName || !productCode || !productTitle) {
        console.log(`⚠️  Skipping row ${i + 1}: Missing required data`);
        console.log(`   SubCategory: "${subCategoryName}", Code: "${productCode}", Title: "${productTitle}"`);
        continue;
      }
      
      console.log(`\n🔄 Processing row ${i + 1}: ${productTitle} (${productCode}) -> ${subCategoryName}`);
      
      try {
        // Find the subcategory
        const subCategory = await getSubCategoryByName(subCategoryName);
        if (!subCategory) {
          console.log(`❌ Subcategory not found: ${subCategoryName}. Please create it first.`);
          continue;
        }
        
        console.log(`✅ Found subcategory: ${subCategoryName} (${subCategory.documentId})`);
        
        // Prepare product data
        const productData = {
          name: productTitle,
          code: productCode,
          size: productSize || ''
        };
        
        // Check if product already exists
        const existingProduct = await getProductByCode(productCode);
        
        if (existingProduct) {
          console.log(`ℹ️  Product already exists: ${productTitle} (${productCode})`);
          // Update existing product
          await updateProduct(existingProduct.documentId, productData, subCategory.documentId);
        } else {
          // Create new product
          await createProduct(productData, subCategory.documentId);
        }
        
        // Add a small delay to avoid overwhelming the API
        await new Promise(resolve => setTimeout(resolve, 200));
        
      } catch (error) {
        console.error(`❌ Error processing row ${i + 1}:`, error.message);
        // Continue with next row instead of stopping
        continue;
      }
    }
    
    console.log('\n🎉 Product upload completed successfully!');
    
  } catch (error) {
    console.error('❌ Error processing products file:', error);
  }
}

// Function to test connection and list some subcategories
async function testConnectionAndListSubcategories() {
  try {
    console.log('🔄 Testing Strapi connection...');
    const response = await makeRequest('GET', '/sub-categories?pagination[limit]=5');
    console.log('✅ Connection successful!');
    
    console.log('\n📋 Available subcategories:');
    response.data.forEach((subCat, index) => {
      console.log(`  ${index + 1}. ${subCat.name} (${subCat.documentId})`);
    });
    
    return true;
  } catch (error) {
    console.error('❌ Connection failed:', error.message);
    return false;
  }
}

// Function to get statistics
async function getUploadStatistics() {
  try {
    const productsResponse = await makeRequest('GET', '/products?pagination[limit]=1');
    const subCategoriesResponse = await makeRequest('GET', '/sub-categories?pagination[limit]=1');
    
    console.log('\n📊 Current Statistics:');
    console.log(`   Products: ${productsResponse.meta.pagination.total}`);
    console.log(`   Subcategories: ${subCategoriesResponse.meta.pagination.total}`);
    
  } catch (error) {
    console.error('❌ Error getting statistics:', error.message);
  }
}

// Run the script
async function main() {
  console.log('🚀 Starting Product upload script...');
  
  // Test connection and show available subcategories
  const connected = await testConnectionAndListSubcategories();
  if (!connected) {
    console.log('❌ Please check your Strapi URL and API token configuration.');
    process.exit(1);
  }
  
  // Show current statistics
  await getUploadStatistics();
  
  // Process the Excel file
  await processProductsFile();
  
  // Show final statistics
  console.log('\n📈 Final Statistics:');
  await getUploadStatistics();
}

// Execute the main function
main().catch(console.error);

// Export functions for testing
module.exports = {
  createSlug,
  generateUniqueSlug,
  getSubCategoryByName,
  getProductByCode,
  createProduct,
  updateProduct,
  processProductsFile
};