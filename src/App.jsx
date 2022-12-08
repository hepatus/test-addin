import { useEffect, useState } from "react";
const Office = window.Office;

function App() {
  const [initialized, setInitialized] = useState(false);
  const [locations, setLocations] = useState([]);

  const updateLocations = () => {
    const item = Office.context.mailbox.item;
    item.enhancedLocation.getAsync((res) => {
      setLocations((res.value || []).map((l) => l.emailAddress));
    });
  };

  const initOffice = async () => {
    Office.onReady(() => {
      updateLocations();

      setInitialized(true);
      const item = Office.context.mailbox.item;
      item.addHandlerAsync(Office.EventType.EnhancedLocationsChanged, () => {
        updateLocations();
      });
    });
  };

  useEffect(() => {
    initOffice();
  }, []);

  if (!initialized) return <div>Office is not initialized</div>;
  return (
    <div>
      <p>Ready</p>
      <p>Locations are : {JSON.stringify(locations)}</p>
    </div>
  );
}

export default App;
