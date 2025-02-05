function main()
 script_id = reaper.NamedCommandLookup("_RS6a0f0f385ca927685554dc1542e587ebbc77aa74")
 reaper.Main_OnCommand(script_id, 0)
reaper.defer(main)
 end
 
 -- Start the script
 main()
