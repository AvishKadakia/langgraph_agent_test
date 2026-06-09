# Agent QA Eval Suite — common tasks. Run `make` to see this help.
APP_PORT ?= 8080

.DEFAULT_GOAL := help

help: ## Show this help
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | \
	  awk 'BEGIN{FS=":.*?## "}{printf "  \033[36m%-12s\033[0m %s\n", $$1, $$2}'

setup: ## First-time setup: create .env from the example
	@test -f .env || (cp .env.example .env && echo "Created .env — open it and fill in any blanks.")
	@touch .env.cookie

login: ## Log in to Azure (needed once; the app reuses this session)
	az login

cookie: ## Capture the web UI session cookie from Chrome into .env.cookie (for agent-web-ui targets)
	@python3 scripts/extract_cookie.py > .env.cookie.tmp 2>/tmp/.cookie_err; \
	if grep -q '^CHAT_API_COOKIE=' .env.cookie.tmp 2>/dev/null; then \
	  mv .env.cookie.tmp .env.cookie; echo "  🍪 web UI session cookie captured"; \
	else \
	  rm -f .env.cookie.tmp; touch .env.cookie; \
	  echo "  ⚠ no session cookie captured — $$(cat /tmp/.cookie_err 2>/dev/null | head -1 | sed 's/^# //')"; \
	  echo "    (only needed for agent-web-ui targets; bearer /chat backends don't need it)"; \
	fi; rm -f /tmp/.cookie_err

up: setup cookie ## Build and start the app (open http://localhost:$(APP_PORT))
	docker compose up --build -d
	@echo ""
	@echo "  ✅ Eval suite running at http://localhost:$(APP_PORT)"
	@echo "  (if 'Not logged in' shows, run: make login)"

down: ## Stop the app
	@touch .env.cookie
	docker compose down

restart: setup cookie ## Re-capture the cookie and restart the app (use after the browser session expires)
	docker compose up -d --force-recreate

logs: ## Follow the app logs
	@touch .env.cookie
	docker compose logs -f

open: ## Open the UI in your browser
	open http://localhost:$(APP_PORT) || xdg-open http://localhost:$(APP_PORT)

ps: ## Show container status
	@touch .env.cookie
	docker compose ps

.PHONY: help setup login cookie up down restart logs open ps
