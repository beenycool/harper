#!/bin/bash

# Harper Word Add-in Build Script

echo "Building Harper Word Add-in..."

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo "Error: Node.js is required but not installed."
    echo "Please install Node.js from https://nodejs.org/"
    exit 1
fi

# Check if pnpm is installed
if ! command -v pnpm &> /dev/null; then
    echo "Error: pnpm is required but not installed."
    echo "Please install pnpm: npm install -g pnpm"
    exit 1
fi

# Install dependencies
echo "Installing dependencies..."
pnpm install

# Build harper.js first (required dependency)
echo "Building harper.js..."
cd ../harper.js
pnpm run build
cd ../word-addon

# Build the add-in
echo "Building Word Add-in..."
pnpm run build

echo "Build complete! Files are in the dist/ directory."
echo ""
echo "To run in development mode:"
echo "1. Run: pnpm run dev-server"
echo "2. In another terminal: pnpm run sideload"
echo ""
echo "For production deployment:"
echo "1. Upload the dist/ folder contents to your web server"
echo "2. Update manifest.xml with your production URLs"
echo "3. Install the add-in using the updated manifest"