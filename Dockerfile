# Use an official R base image
FROM r-base:4.3.1

# Install system dependencies for R packages
RUN apt-get update && apt-get install -y \
    libcurl4-openssl-dev \
    libssl-dev \
    libxml2-dev \
    libfontconfig1-dev \
    libfreetype6-dev \
    libpng-dev \
    libtiff5-dev \
    libjpeg-dev \
    zlib1g-dev \
    pandoc \
    curl \
    git \
    && rm -rf /var/lib/apt/lists/*

# Install required R packages
RUN R -e "install.packages(c('plumber', 'readxl', 'openxlsx', 'writexl', 'dplyr', 'tidyr'), repos='https://cloud.r-project.org/')"

# Copy API code into the container
WORKDIR /app
COPY . /app

# Expose port 8000 for Plumber
EXPOSE 8000

# Run the API
CMD ["R", "-e", "pr <- plumber::plumb('plumber.R'); pr$run(host = '0.0.0.0', port = 8000)"]
