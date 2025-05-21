.PHONY:  help flake8 test migrate mysql mysql shell bash debug rebuild publish docs

publish:
	hatch build
	hatch publish

update_snapshot:
	uv run pytest --update-snapshot

docs: ## bash
	uv run mkdocs serve

build_docs:
	uv run mkdocs build

deploy_docs:
	uv run mkdocs gh-deploy -r upstream

