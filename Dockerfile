FROM node:20

# Install LibreOffice full CLI with all required dependencies
RUN apt-get update && \
    apt-get install -y libreoffice libreoffice-core libreoffice-common libreoffice-writer libreoffice-calc libreoffice-impress libreoffice-draw fonts-dejavu fonts-liberation && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install

# Copy project files
COPY . .

# Build the application (if using TypeScript)
RUN npm run build

# Make output folders
RUN mkdir -p /app/output-generated/excel /app/output-generated/pdf

# Start the app
CMD ["node", "dist/app.js"]
