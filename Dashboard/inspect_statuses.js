import { fetchAllDataAcrossBase } from './src/api/airtable.js';
import dotenv from 'dotenv';
dotenv.config({ path: '.ENV' });

// Mock import.meta.env for Node.js
if (!global.import) global.import = {};
if (!global.import.meta) global.import.meta = {};
global.import.meta.env = {
    VITE_AIRTABLE_TOKEN: process.env.VITE_AIRTABLE_TOKEN,
    VITE_AIRTABLE_BASE_ID: process.env.VITE_AIRTABLE_BASE_ID,
    VITE_AIRTABLE_TABLE_NAME: process.env.VITE_AIRTABLE_TABLE_NAME,
};

async function inspectStatuses() {
    console.log('Fetching all data to inspect statuses...');
    try {
        const data = await fetchAllDataAcrossBase();
        const statusCounts = {};
        let totalRecords = 0;
        let recordsWithStatus = 0;

        for (const [tableName, records] of Object.entries(data)) {
            totalRecords += records.length;
            records.forEach(r => {
                const status = r.fields['Registration Status'];
                if (status) {
                    recordsWithStatus++;
                    statusCounts[status] = (statusCounts[status] || 0) + 1;
                } else {
                    statusCounts['[EMPTY]'] = (statusCounts['[EMPTY]'] || 0) + 1;
                }
            });
        }

        console.log('Total Records:', totalRecords);
        console.log('Records with any status:', recordsWithStatus);
        console.log('Status Distribution:');
        console.table(statusCounts);
    } catch (error) {
        console.error('Error during inspection:', error);
    }
}

inspectStatuses();
