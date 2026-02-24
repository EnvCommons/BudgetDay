# BudgetDay Environment

OpenReward environment for Budget Day tasks.

## Status

This is a stub implementation with all boilerplate in place.

## Structure

- `budgetday.py` - Main environment class (extends CLIEnvironment)
- `cli_environment.py` - Base class providing CLI tools (bash, read, write, grep, glob)
- `server.py` - Minimal server wrapper
- `test_agent.py` - Local agent testing
- `constants.py` - Path configuration for local/production
- `requirements.txt` - Python dependencies
- `Dockerfile` - Container configuration

## TODO

- [ ] Define actual tasks in `list_tasks()`
- [ ] Implement `get_prompt()` with real prompts
- [ ] Implement `submit_answer()` validation logic
- [ ] Add data files (if needed)
- [ ] Implement CLI tools (bash, read, write, etc.) if sandbox needed
- [ ] Add tests
- [ ] Create DATA_UPLOAD.md if external data required

## Local Development

```bash
# Install dependencies
uv pip install -r requirements.txt

# Run server
python server.py

# Test with agent (in another terminal)
export OPENAI_API_KEY=your_key
python test_agent.py
```

## Docker

```bash
# Build
docker build -t budgetday:test .

# Run
docker run -p 8080:8080 budgetday:test
```
