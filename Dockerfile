# Start from a Go image
FROM golang:1.25.4-alpine3.22

# Setup Work Directory
WORKDIR /app

# Copy Go modules and install dependencies
COPY go.mod go.sum ./
RUN go mod download

# Copy the rest of the application code
COPY . .

# Build the Go application
RUN go build -o golangldapservice .

# Expose the port your app will run on
EXPOSE ${PORT}

ENV LDAP_IP=10.144.1.6
ENV LDAP_DNS=psth.com
ENV SERVER_PORT=8090

ENV SQL_USER=${SQL_USER}
ENV SQL_PASSWORD=${SQL_PASSWORD}
ENV SQL_DB=DB_EDI
ENV SQL_HOST=10.144.1.101
ENV SQL_PORT=1433

ENV SERVER_PORT=8090

# Run the application
CMD ["./golangldapservice"]