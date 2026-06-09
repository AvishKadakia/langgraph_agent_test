# Agent QA Eval Suite — common tasks. Run `make` to see this help.
# (For non-developers without make/docker, just double-click run.command (macOS)
#  or run.bat (Windows) — see README.)
APP_PORT ?= 8080

.DEFAULT_GOAL := help

help: ## Show this help
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | \
	  awk 'BEGIN{FS=":.*?## "}{printf "  \033[36m%-12s\033[0m %s\n", $$1, $$2}'

setup: ## First-time setup: create .env from the example
	@test -f .env || (cp .env.example .env && echo "Created .env — open it and fill in any blanks.")

login: ## Log in to Azure (needed once; the app reuses this session)
	az login

up: setup ## Build and start the app (open http://localhost:$(APP_PORT))
	docker compose up --build -d
	@echo ""
	@echo "  ✅ Eval suite running at http://localhost:$(APP_PORT)"
	@echo "  (if 'Not logged in' shows, run: make login)"

down: ## Stop the app
	docker compose down

restart: ## Restart the app
	docker compose restart

logs: ## Follow the app logs
	docker compose logs -f

open: ## Open the UI in your browser
	open http://localhost:$(APP_PORT) || xdg-open http://localhost:$(APP_PORT)

ps: ## Show container status
	docker compose ps

.PHONY: help setup login up down restart logs open ps
