#!/bin/bash

echo "Starting MacroFlow AI..."

# Check if frontend dependencies need to be installed
echo "Checking frontend dependencies..."
cd frontend
if [ ! -d "node_modules" ]; then
    echo "Installing frontend dependencies..."
    npm install
fi

# Start Flask backend in background
echo "Starting Flask backend (port 5000)..."
cd ../backend
python app.py &
BACKEND_PID=$!

# Wait for backend to start
sleep 3

# Start frontend server in background
echo "Starting frontend server (port 3000)..."
cd ../frontend
node server.js &
FRONTEND_PID=$!

# Wait for frontend to start
sleep 2

# Start Office add-in debugging
echo "Starting Office add-in debugging..."
npx office-addin-debugging start manifest.xml &
ADDIN_PID=$!

echo ""
echo "MacroFlow AI is starting up..."
echo "Backend: http://localhost:5000"
echo "Frontend: https://localhost:3000"
echo ""
echo "Press Ctrl+C to stop all servers."
echo ""

# Function to cleanup background processes
cleanup() {
    echo "Stopping MacroFlow AI..."
    kill $BACKEND_PID $FRONTEND_PID $ADDIN_PID 2>/dev/null
    exit 0
}

# Set trap to cleanup on script exit
trap cleanup SIGINT

# Wait for any background job to finish
wait