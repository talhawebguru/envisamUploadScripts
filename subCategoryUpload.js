const XLSX = require('xlsx');
const axios = require('axios');

// Configuration
const STRAPI_URL = 'http://localhost:1337'; // Change this to your Strapi URL
const STRAPI_TOKEN = '61203c16fb1450a1b21c285e797de58ba13106cc1cbed92d2e6bc67c4b3d17ff18659d5ec9094e3fc944cf7bdce5d1faddaaa222f3f75de72d93643f56564b9359a7ff02cb0a0ea627eba4a2b7030948198b1e53751d1221bab4b5fd281b098ba9e4dec24762c25109b60c4bc9ed98a1f856d69fd89d051cb8b81152785cd6c9'; // Replace with your actual API token
const EXCEL_FILE_PATH = './SubCategoryUploads.xlsx'; // Replace with your Excel file path


// Helper function to create slug from name
function createSlug(name) {
  return name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
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

// Function to check if category exists
async function getCategoryByName(name) {
  try {
    const response = await makeRequest('GET', `/categories?filters[name][$eq]=${encodeURIComponent(name)}`);
    return response.data.length > 0 ? response.data[0] : null;
  } catch (error) {
    console.error('Error checking category:', error);
    return null;
  }
}

// Function to check if subcategory exists
async function getSubCategoryByName(name) {
  try {
    const response = await makeRequest('GET', `/sub-categories?filters[name][$eq]=${encodeURIComponent(name)}`);
    return response.data.length > 0 ? response.data[0] : null;
  } catch (error) {
    console.error('Error checking subcategory:', error);
    return null;
  }
}

// Function to create category
async function createCategory(name) {
  const categoryData = {
    data: {
      name: name,
      slug: createSlug(name),
      publishedAt: new Date().toISOString(), // Auto-publish
    }
  };

  try {
    const response = await makeRequest('POST', '/categories', categoryData);
    console.log(`✅ Created category: ${name}`);
    return response.data;
  } catch (error) {
    console.error(`❌ Failed to create category: ${name}`, error);
    throw error;
  }
}

// Function to create subcategory
async function createSubCategory(name, categoryDocumentId) {
  const subCategoryData = {
    data: {
      name: name,
      slug: createSlug(name),
      publishedAt: new Date().toISOString(), // Auto-publish
    }
  };

  try {
    const response = await makeRequest('POST', '/sub-categories', subCategoryData);
    console.log(`✅ Created subcategory: ${name}`);
    return response.data;
  } catch (error) {
    console.error(`❌ Failed to create subcategory: ${name}`, error);
    throw error;
  }
}

// Function to update category with subcategory relationship
async function updateCategoryWithSubCategory(categoryDocumentId, subCategoryDocumentId) {
  try {
    // First, get the current category with its subcategories
    const currentCategory = await makeRequest('GET', `/categories/${categoryDocumentId}?populate=sub_categories`);
    
    // Get current subcategory document IDs - handle both v4 and v5 response structures
    let currentSubCategoryIds = [];
    if (currentCategory.data.sub_categories) {
      if (Array.isArray(currentCategory.data.sub_categories)) {
        // Direct array of subcategories
        currentSubCategoryIds = currentCategory.data.sub_categories.map(sub => sub.documentId);
      } else if (currentCategory.data.sub_categories.data) {
        // Wrapped in data object (v4 style)
        currentSubCategoryIds = currentCategory.data.sub_categories.data.map(sub => sub.documentId);
      }
    }
    
    // Add the new subcategory document ID if it's not already there
    if (!currentSubCategoryIds.includes(subCategoryDocumentId)) {
      currentSubCategoryIds.push(subCategoryDocumentId);
    }

    // Update the category with the new subcategory relationship
    const updateData = {
      data: {
        sub_categories: currentSubCategoryIds
      }
    };

    await makeRequest('PUT', `/categories/${categoryDocumentId}`, updateData);
    console.log(`✅ Updated category relationship with subcategory`);
  } catch (error) {
    console.error(`❌ Failed to update category relationship:`, error);
    throw error;
  }
}

// Main function to process Excel file
async function processExcelFile() {
  try {
    console.log('📖 Reading Excel file...');
    
    // Read the Excel file
    const workbook = XLSX.readFile(EXCEL_FILE_PATH);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert to JSON with header row
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    console.log(`📊 Found ${data.length} rows in Excel file`);
    
    // Skip header row and process data rows
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Using column positions: 0=S#, 1=Main Cat, 2=Sub Cat
      const categoryName = row[1]?.toString().trim();
      const subCategoryName = row[2]?.toString().trim();
      
      if (!categoryName || !subCategoryName) {
        console.log(`⚠️  Skipping row ${i + 1}: Missing category or subcategory name`);
        console.log(`   Category: "${categoryName}", Subcategory: "${subCategoryName}"`);
        continue;
      }
      
      console.log(`\n🔄 Processing row ${i + 1}: ${categoryName} -> ${subCategoryName}`);
      
      try {
        // Check if category exists, create if not
        let category = await getCategoryByName(categoryName);
        if (!category) {
          category = await createCategory(categoryName);
        } else {
          console.log(`ℹ️  Category already exists: ${categoryName}`);
        }
        
        // Check if subcategory exists, create if not
        let subCategory = await getSubCategoryByName(subCategoryName);
        if (!subCategory) {
          subCategory = await createSubCategory(subCategoryName, category.documentId);
        } else {
          console.log(`ℹ️  Subcategory already exists: ${subCategoryName}`);
        }
        
        // Update category with subcategory relationship
        await updateCategoryWithSubCategory(category.documentId, subCategory.documentId);
        
        // Add a small delay to avoid overwhelming the API
        await new Promise(resolve => setTimeout(resolve, 100));
        
      } catch (error) {
        console.error(`❌ Error processing row ${i + 1}:`, error.message);
        // Continue with next row instead of stopping
        continue;
      }
    }
    
    console.log('\n🎉 Upload completed successfully!');
    
  } catch (error) {
    console.error('❌ Error processing Excel file:', error);
  }
}

// Function to test connection
async function testConnection() {
  try {
    console.log('🔄 Testing Strapi connection...');
    await makeRequest('GET', '/categories?pagination[limit]=1');
    console.log('✅ Connection successful!');
    return true;
  } catch (error) {
    console.error('❌ Connection failed:', error.message);
    return false;
  }
}

// Run the script
async function main() {
  console.log('🚀 Starting Strapi upload script...');
  
  // Test connection first
  const connected = await testConnection();
  if (!connected) {
    console.log('❌ Please check your Strapi URL and API token configuration.');
    process.exit(1);
  }
  
  // Process the Excel file
  await processExcelFile();
}

// Execute the main function
main().catch(console.error);

// Export functions for testing
module.exports = {
  createSlug,
  getCategoryByName,
  getSubCategoryByName,
  createCategory,
  createSubCategory,
  updateCategoryWithSubCategory,
  processExcelFile
};