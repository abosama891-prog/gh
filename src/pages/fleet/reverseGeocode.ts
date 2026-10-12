type Coordinates = { latitude: number; longitude: number };

const addressPart = (value: unknown) => typeof value === 'string' ? value.trim() : '';

const formatAddress = (parts: unknown[]) => [...new Set(parts.map(addressPart).filter(Boolean))].join('، ');

const readNominatimAddress = (result: unknown) => {
  if (!result || typeof result !== 'object') return '';
  const place = result as {
    name?: unknown;
    display_name?: unknown;
    address?: Record<string, unknown>;
  };
  const address = place.address ?? {};
  return formatAddress([
    place.name,
    address.road,
    address.pedestrian,
    address.neighbourhood,
    address.suburb,
    address.village,
    address.town,
    address.city,
    address.county,
    address.state,
    address.country,
  ]) || addressPart(place.display_name);
};

const readPhotonAddress = (result: unknown) => {
  if (!result || typeof result !== 'object') return '';
  const features = (result as { features?: unknown }).features;
  if (!Array.isArray(features) || !features[0] || typeof features[0] !== 'object') return '';
  const properties = (features[0] as { properties?: Record<string, unknown> }).properties ?? {};
  return formatAddress([
    properties.name,
    properties.street,
    properties.district,
    properties.suburb,
    properties.city,
    properties.county,
    properties.state,
    properties.country,
  ]);
};

const requestAddress = async (url: string, readAddress: (result: unknown) => string) => {
  const response = await fetch(url, {
    headers: { Accept: 'application/json' },
    signal: AbortSignal.timeout(10000),
  });
  if (!response.ok) throw new Error(`Address lookup failed with HTTP ${response.status}`);
  const address = readAddress(await response.json());
  if (!address) throw new Error('Address lookup returned no readable place name');
  return address;
};

export const reverseGeocode = async ({ latitude, longitude }: Coordinates): Promise<string> => {
  const nominatimParams = new URLSearchParams({
    format: 'jsonv2',
    lat: String(latitude),
    lon: String(longitude),
    zoom: '18',
    addressdetails: '1',
    'accept-language': 'ar',
  });
  try {
    return await requestAddress(
      `https://nominatim.openstreetmap.org/reverse?${nominatimParams}`,
      readNominatimAddress,
    );
  } catch (nominatimError) {
    const photonParams = new URLSearchParams({
      lat: String(latitude),
      lon: String(longitude),
      lang: 'ar',
    });
    try {
      return await requestAddress(
        `https://photon.komoot.io/reverse?${photonParams}`,
        readPhotonAddress,
      );
    } catch (photonError) {
      console.error('All reverse geocoding providers failed', { nominatimError, photonError });
      throw new Error('تعذر العثور على اسم عنوان لهذا الموقع');
    }
  }
};
