publish:
	poetry build
	poetry publish

update_snapshot:
	pytest --update-snapshot
