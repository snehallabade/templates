FROM node:20

# Install LibreOffice and required dependencies
RUN apt-get update && \
    apt-get install -y libreoffice && \
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

# Build the application
RUN npm run build

# Create necessary directories for file generation
RUN mkdir -p /app/output-generated/excel /app/output-generated/pdf

# Start the app
CMD ["node", "dist/app.js"] 