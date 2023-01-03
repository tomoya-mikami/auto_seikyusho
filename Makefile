.PHONY: build
build:
	docker build . -t auto_seikyusho

.PHONEY: run
run:
	docker run -v $(shell pwd):/app auto_seikyusho
