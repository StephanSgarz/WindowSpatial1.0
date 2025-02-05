-- Get the master track
local targetProjectName = "ambisonic_4.RPP"
local currentProjectName = reaper.GetProjectName(0, "")
if currentProjectName == targetProjectName then
  local track = reaper.GetMasterTrack(0)

  -- Find the SPARTA Binauraliser FX (assuming it's the first FX, index 0)
  local fx_index = 0

  -- Parameter indices for azimuth and elevation
  local azimuth_param_index = 9 -- Adjust if azimuth is on a different parameter index
  local elevation_param_index = 10 -- Adjust if elevation is on a different parameter index

  -- New values for azimuth and elevation
  local azimuth_value = 0.7
  local elevation_value = 0.8

  -- Set the azimuth parameter
  reaper.TrackFX_SetParam(track, fx_index, azimuth_param_index, azimuth_value)

  -- Set the elevation parameter
  reaper.TrackFX_SetParam(track, fx_index, elevation_param_index, elevation_value)

  reaper.UpdateArrange() -- Update the arrangement view to reflect changes
  end
