import { useState } from 'react';
import { MapContainer, Marker, TileLayer, useMapEvents } from 'react-leaflet';
import L, { type LatLngExpression } from 'leaflet';
import 'leaflet/dist/leaflet.css';
import { reverseGeocode } from './reverseGeocode';

type Coordinate = { latitude: number; longitude: number };
type RouteMapPickerProps = {
  initialCoordinate?: Coordinate;
  onCancel: () => void;
  onConfirm: (address: string) => void;
};

const defaultCenter: Coordinate = { latitude: 30.0444, longitude: 31.2357 };
const markerIcon = L.divIcon({
  className: 'route-map-marker-icon',
  html: '<span></span>',
  iconSize: [26, 34],
  iconAnchor: [13, 32],
});

const SelectMapPoint = ({ onSelect }: { onSelect: (coordinate: Coordinate) => void }) => {
  useMapEvents({
    click: ({ latlng }) => onSelect({ latitude: latlng.lat, longitude: latlng.lng }),
  });
  return null;
};

const RouteMapPicker = ({ initialCoordinate, onCancel, onConfirm }: RouteMapPickerProps) => {
  const [coordinate, setCoordinate] = useState<Coordinate | null>(initialCoordinate ?? null);
  const [isResolvingAddress, setIsResolvingAddress] = useState(false);
  const [addressError, setAddressError] = useState('');
  const center: LatLngExpression = initialCoordinate
    ? [initialCoordinate.latitude, initialCoordinate.longitude]
    : [defaultCenter.latitude, defaultCenter.longitude];

  const resolveAddress = async (allowCoordinates = false) => {
    if (!coordinate || isResolvingAddress) return;
    if (allowCoordinates) {
      onConfirm(`GPS: ${coordinate.latitude.toFixed(6)}, ${coordinate.longitude.toFixed(6)}`);
      return;
    }

    setIsResolvingAddress(true);
    setAddressError('');
    try {
      onConfirm(await reverseGeocode(coordinate));
    } catch (error) {
      console.error('Could not resolve selected map point to an address', error);
      setAddressError('تعذر جلب اسم العنوان. تحقق من الاتصال بالإنترنت أو استخدم الإحداثيات بدلًا منه.');
    } finally {
      setIsResolvingAddress(false);
    }
  };

  return (
    <div className="route-map-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) onCancel(); }}>
      <section className="route-map-dialog" role="dialog" aria-modal="true" aria-labelledby="route-map-title" dir="rtl">
        <header className="route-map-header">
          <div><h2 id="route-map-title">اختيار الموقع من الخريطة</h2><p>اضغط على المكان المطلوب لتحديده</p></div>
          <button type="button" className="route-map-close" onClick={onCancel} aria-label="إغلاق الخريطة">×</button>
        </header>
        <div className="route-map-canvas">
          <MapContainer center={center} zoom={initialCoordinate ? 14 : 6} scrollWheelZoom>
            <TileLayer
              attribution='&copy; <a href="https://www.openstreetmap.org/copyright" target="_blank" rel="noreferrer">OpenStreetMap</a> contributors'
              url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
            />
            <SelectMapPoint onSelect={setCoordinate} />
            {coordinate && <Marker position={[coordinate.latitude, coordinate.longitude]} icon={markerIcon} />}
          </MapContainer>
        </div>
        <footer className="route-map-footer">
          <div className="route-map-selection">
            <span>{coordinate ? `الموقع المحدد: ${coordinate.latitude.toFixed(6)}, ${coordinate.longitude.toFixed(6)}` : 'لم يتم اختيار موقع بعد'}</span>
            {addressError && <strong role="alert">{addressError}</strong>}
          </div>
          <div className="route-map-actions">
            <button type="button" className="fleet-button fleet-button-light" onClick={onCancel} disabled={isResolvingAddress}>إلغاء</button>
            {addressError && <button type="button" className="fleet-button fleet-button-light" onClick={() => void resolveAddress(true)}>استخدام الإحداثيات</button>}
            <button type="button" className="fleet-button fleet-button-primary" disabled={!coordinate || isResolvingAddress} onClick={() => void resolveAddress()}>{isResolvingAddress ? 'جارٍ جلب العنوان...' : 'اعتماد العنوان'}</button>
          </div>
        </footer>
      </section>
    </div>
  );
};

export default RouteMapPicker;
