build:
	docker compose build

tests: build
	docker compose run --rm web-api pytest
