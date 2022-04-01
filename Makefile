TEST?=$$(go list ./... | grep -v 'vendor')
HOSTNAME=sbl.systems
NAMESPACE=synwork
NAME=excel
BINARY=synwork-processor-${NAME}
VERSION=0.0.1
OS_ARCH=linux_amd64

default: install

build:
	CGO_ENABLED=0 go build -o ${BINARY}

release:
	CGO_ENABLED=0 GOOS=darwin GOARCH=amd64 go build -o ./bin/${BINARY}_${VERSION}_darwin_amd64
	CGO_ENABLED=0 GOOS=freebsd GOARCH=386 go build -o ./bin/${BINARY}_${VERSION}_freebsd_386
	CGO_ENABLED=0 GOOS=freebsd GOARCH=amd64 go build -o ./bin/${BINARY}_${VERSION}_freebsd_amd64
	CGO_ENABLED=0 GOOS=freebsd GOARCH=arm go build -o ./bin/${BINARY}_${VERSION}_freebsd_arm
	CGO_ENABLED=0 GOOS=linux GOARCH=386 go build -o ./bin/${BINARY}_${VERSION}_linux_386
	CGO_ENABLED=0 GOOS=linux GOARCH=amd64 go build -o ./bin/${BINARY}_${VERSION}_linux_amd64
	CGO_ENABLED=0 GOOS=linux GOARCH=arm go build -o ./bin/${BINARY}_${VERSION}_linux_arm
	CGO_ENABLED=0 GOOS=openbsd GOARCH=386 go build -o ./bin/${BINARY}_${VERSION}_openbsd_386
	CGO_ENABLED=0 GOOS=openbsd GOARCH=amd64 go build -o ./bin/${BINARY}_${VERSION}_openbsd_amd64
	CGO_ENABLED=0 GOOS=solaris GOARCH=amd64 go build -o ./bin/${BINARY}_${VERSION}_solaris_amd64
	CGO_ENABLED=0 GOOS=windows GOARCH=386 go build -o ./bin/${BINARY}_${VERSION}_windows_386
	CGO_ENABLED=0 GOOS=windows GOARCH=amd64 go build -o ./bin/${BINARY}_${VERSION}_windows_amd64

install: build
	mkdir -p ~/.synwork.d/plugins/${HOSTNAME}/${NAMESPACE}/${NAME}/${VERSION}/${OS_ARCH}
	mv ${BINARY} ~/.synwork.d/plugins/${HOSTNAME}/${NAMESPACE}/${NAME}/${VERSION}/${OS_ARCH}

test: 
	go test -i $(TEST) || exit 1                                                   
	echo $(TEST) | xargs -t -n4 go test $(TESTARGS) -timeout=30s -parallel=4                    

testacc: 
	TF_ACC=1 go test $(TEST) -v $(TESTARGS) -timeout 120m   