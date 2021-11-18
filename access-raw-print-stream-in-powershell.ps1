# This example demonstrates how to access the raw print stream in Print Distributor using PowerShell

# Reference the Model assembly in the Print Distributor installation directory
Add-Type -path 'C:\Program Files\Print Distributor\Model.dll'

# Retrieve the original stream
$stream = $context.GetValue("_stream")

# Create a new stream to store the modified print stream
# The HybridStream inherits from System.IO.Stream and
# automatically switches to a file based stream when the
# data gets over 10MB.
$newStream = New-Object -TypeName Model.HybridStream

# Do something with the stream
# All actions in Print Distributor seek the start of the stream if they need to use it
$stream.Seek(0, 0) | Out-null

$stream.CopyTo($newStream)
# Truncate the stream after 1000 bytes
$newStream.SetLength(1000) 

# Replace the original
$newStream.Seek(0, 0) | Out-null
# Wipe out the original data
$stream.SetLength(0)
# Copy the new version to the original stream
$newStream.CopyTo($stream)

# Don't close the original, only what we have created
$newStream.Close()
