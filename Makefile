.PHONY:  help flake8 test migrate mysql mysql shell bash debug rebuild publish docs

publish:
	poetry build
	poetry publish

update_snapshot:
	pytest --update-snapshot

docs: ## bash
	poetry run mkdocs serve

build_docs:
	poetry run mkdocs build

deploy_docs:
	poetry run mkdocs gh-deploy -r upstream

