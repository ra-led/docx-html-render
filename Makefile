build:
	docker compose build

tests: build
	docker compose run --rm docx-to-html pytest
