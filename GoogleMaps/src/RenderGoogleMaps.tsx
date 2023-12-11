import React from "react";
import { IInputs } from "./generated/ManifestTypes";
import { Spinner } from "@fluentui/react";
import { GoogleMap, Marker, useJsApiLoader } from "@react-google-maps/api";

const DEFAULT_POSITION = {
  lat: 56.05411222396458,
  lng: 14.623579711215468,
};

export function RenderGoogleMaps({
  context,
  updatePosition,
}: {
  context: ComponentFramework.Context<IInputs>;
  // eslint-disable-next-line no-unused-vars
  updatePosition: (position: string | undefined) => void;
}) {
  const [center, setCenter] = React.useState(DEFAULT_POSITION);

  /**
   * Split the position into its component parts
   * @returns {lat: number, lng: number}
   */
  const splitPosition = () => {
    const position = context.parameters.position.raw;
    if (position) {
      const split = position.split(":");
      return { lat: parseFloat(split[1]), lng: parseFloat(split[2]) };
    }
    return DEFAULT_POSITION;
  };

  const { isLoaded } = useJsApiLoader({
    id: "google-map-script",
    googleMapsApiKey: "AIzaSyDz9hlb0-MzJnu8XWwNeTExGLUNithdBFI",
  });

  /**
   * UseEffect to run if address lines change
   * Will geocode the address and update the position attribute
   */
  React.useEffect(() => {
    if (isLoaded) {
      const combinedAddress = `${context.parameters.addressLine1.raw} ${context.parameters.addressLine2.raw} ${context.parameters.addressLine3.raw}`;
      const geocoder = new window.google.maps.Geocoder();
      geocoder.geocode({ address: combinedAddress }, function (results) {
        if (results)
          updatePosition(
            `${
              results[0].plus_code?.global_code
            }:${results[0].geometry?.location?.lat()}:${results[0].geometry?.location?.lng()}`
          );
      });
    }
  }, [
    isLoaded,
    context.parameters.addressLine1.raw,
    context.parameters.addressLine2.raw,
    context.parameters.addressLine3.raw,
  ]);

  /**
   * UseEffect to run if position attribute changes
   * Will update the center of the map
   */
  React.useEffect(() => {
    if (
      isLoaded &&
      context.parameters.position.raw &&
      context.parameters.position.raw.length > 0
    ) {
      setCenter(splitPosition());
    } else setCenter(DEFAULT_POSITION);
  }, [isLoaded, context.parameters.position.raw]);

  return !isLoaded ? (
    <Spinner label="Loading..." labelPosition="right" />
  ) : (
    <GoogleMap
      mapContainerStyle={{
        width: `${context.mode.allocatedWidth}px`,
        height: "400px",
      }}
      zoom={15}
      center={center}
      options={{ fullscreenControl: false, streetViewControl: false }}
    >
      <Marker position={center} />
    </GoogleMap>
  );
}
