const AVPN_SUMMARY_ENDPOINT = '/api/thinkific/avpn-summary';

export async function fetchAvpnSummary(options = {}) {
    const { forceRefresh = false } = options;
    const url = new URL(AVPN_SUMMARY_ENDPOINT, window.location.origin);
    if (forceRefresh) {
        url.searchParams.set('refresh', '1');
    }

    const response = await fetch(url.toString(), {
        headers: {
            Accept: 'application/json',
        },
    });

    if (!response.ok) {
        const body = await response.text();
        throw new Error(`Failed to load AVPN summary (${response.status}): ${body || 'Unknown error'}`);
    }

    const data = await response.json();
    if (!Array.isArray(data?.rows)) {
        throw new Error('AVPN summary response is invalid.');
    }

    return data;
}
