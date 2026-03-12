const TOKEN = 'patn8rdQYk2z93WQ4.8b3957ebdcf3855fd11e887f430546674bbecbf61ed1feb00d42beacbdf275f4';
const BASE_ID = 'appHneDvHaRJVWgk7';

async function inspect() {
    try {
        const metaUrl = `https://api.airtable.com/v0/meta/bases/${BASE_ID}/tables`;
        const metaRes = await fetch(metaUrl, { headers: { Authorization: `Bearer ${TOKEN}` } });
        const metaData = await metaRes.json();

        const statusCounts = {};
        let totalRecords = 0;

        console.log(`Checking ${metaData.tables.length} tables...`);

        for (const table of metaData.tables) {
            let offset = null;
            do {
                const url = new URL(`https://api.airtable.com/v0/${BASE_ID}/${table.id}`);
                if (offset) url.searchParams.set('offset', offset);

                const res = await fetch(url.toString(), { headers: { Authorization: `Bearer ${TOKEN}` } });
                const data = await res.json();

                totalRecords += data.records.length;
                data.records.forEach(r => {
                    const status = r.fields['Registration Status'];
                    const val = status || '[EMPTY]';
                    statusCounts[val] = (statusCounts[val] || 0) + 1;
                });

                offset = data.offset;
            } while (offset);
        }

        console.log('Total Records:', totalRecords);
        console.log('Status Distribution:');
        Object.entries(statusCounts)
            .sort((a, b) => b[1] - a[1])
            .forEach(([status, count]) => {
                console.log(`${status.padEnd(20)}: ${count}`);
            });

    } catch (e) {
        console.error(e);
    }
}

inspect();
