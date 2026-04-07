### Why the container exited immediately

Your image’s default command is a shell that runs and then exits because there is no long‑running process keeping the container alive. The image metadata shows CMD ["/bin/bash"] so when Docker starts the container it runs /bin/bash non‑interactively, the shell exits immediately, and the container stops with exit code 0.

Quick diagnostics
- See the container exit code and last logs:  
  docker ps -a --filter id=a93ff603efc8 --format "table {{.ID}}\t{{.Image}}\t{{.Status}}" && docker logs a93ff603efc8
- Inspect configured command and entrypoint:  
  docker image inspect jeremiahwong/ppt-agent-tools:latest --format '{{json .Config.Cmd}} {{json .Config.Entrypoint}}'

---

### How to start a container that stays running

Choose one of these approaches depending on whether you want the container to run a real service or just stay alive for interactive access.

- Run a long‑lived command (recommended for quick dev sessions):  
  docker run -d --name ppt-agent-tools \
    --restart unless-stopped \
    jeremiahwong/ppt-agent-tools:latest \
    tail -f /dev/null

- Start with bash in the background (keeps it alive):  
  docker run -d --name ppt-agent-tools \
    jeremiahwong/ppt-agent-tools:latest \
    /bin/bash -c "while true; do sleep 86400; done"

- Update the Dockerfile to provide a dev-friendly default (persists across runs):  
  Replace CMD ["/bin/bash"] with CMD ["tail","-f","/dev/null"] or use your real service command as CMD.

---

### How to open a shell anytime

- If the container is already running, use docker exec for an interactive shell:  
  docker exec -it ppt-agent-tools /bin/bash  
  If bash isn’t present, use /bin/sh.

- Start a new ephemeral interactive container (no background):  
  docker run --rm -it -v "$PWD":/app:delegated --entrypoint /bin/bash jeremiahwong/ppt-agent-tools:latest

- Start container detached then attach shell later:  
  docker run -d --name ppt-agent-tools jeremiahwong/ppt-agent-tools:latest tail -f /dev/null  
  docker exec -it ppt-agent-tools /bin/bash

---

### Useful lifecycle and debug commands

- See running containers: docker ps  
- See all containers and their exit reasons: docker ps -a  
- Inspect exit code: docker inspect --format='{{.State.ExitCode}} {{.State.Status}}' <container-id>  
- View last logs: docker logs <container-id>  
- Start/stop/remove: docker start <name> ; docker stop <name> ; docker rm <name>

---

### Recommended development workflow

1. For development mount your repo and keep container alive:
   docker run -d --name ppt-agent-tools -v "$PWD":/app:delegated -w /app jeremiahwong/ppt-agent-tools:latest tail -f /dev/null

2. Exec into container to test and iterate:
   docker exec -it ppt-agent-tools /bin/bash

3. When you have a real service, set that as CMD in the Dockerfile and use proper process management (tini) and a non‑root user before promoting to production.

---

### Quick checklist to fix your current run

- [ ] Start container with a long‑running command (tail -f /dev/null) or change Dockerfile CMD.  
- [ ] Use docker exec -it <name> /bin/bash to get shell access.  
- [ ] Inspect logs and image Config if unexpected exits persist.  

---

https://copilot.microsoft.com/shares/djU8DncibzmtEfM8bhcgY  
https://copilot.microsoft.com/shares/mNXVMwZQwoudMBnwgWH6D  

---

### Start a container in the background

- Run the image detached with a stable name, restart policy, and an optional bind-mount for live code:
```bash
# simple detached run with name and restart policy
docker run -d --name ppt-agent-tools \
  --restart unless-stopped \
  jeremiahwong/ppt-agent-tools:latest

# with repo-local bind mount so /app in the container reflects your working tree
docker run -d --name ppt-agent-tools \
  --restart unless-stopped \
  -v "$PWD":/app:delegated \
  jeremiahwong/ppt-agent-tools:latest
```

- If your container would normally exit immediately because CMD is a shell, keep it alive by running a long-lived command:
```bash
docker run -d --name ppt-agent-tools \
  --restart unless-stopped \
  -v "$PWD":/app:delegated \
  jeremiahwong/ppt-agent-tools:latest \
  tail -f /dev/null
```

- Resource limits (recommended for production/dev ergonomics):
```bash
docker run -d --name ppt-agent-tools \
  --restart unless-stopped \
  --cpus="1.0" --memory="512m" \
  -v "$PWD":/app:delegated \
  jeremiahwong/ppt-agent-tools:latest \
  tail -f /dev/null
```

---

### Shell access any time (recommended approach)

- Use docker exec to open an interactive shell in a running container:
```bash
# get an interactive bash shell
docker exec -it ppt-agent-tools /bin/bash

# if bash is not present, fall back to sh
docker exec -it ppt-agent-tools /bin/sh
```

- If you started the container without a name, find the container id first:
```bash
docker ps                    # see NAME and CONTAINER ID
docker exec -it <CONTAINER_ID> /bin/bash
```

---

### Alternatives: attach, logs and one-off shells

- Attach to main process STDIN/STDOUT (less flexible; can disrupt process):
```bash
docker attach ppt-agent-tools
# detach from attach session without stopping: Ctrl-p Ctrl-q
```

- Run a one-off shell (starts a new ephemeral container from the image):
```bash
docker run --rm -it -v "$PWD":/app:delegated --entrypoint /bin/bash jeremiahwong/ppt-agent-tools:latest
```

- Inspect logs if you need to debug containers started in background:
```bash
docker logs -f ppt-agent-tools
docker logs --since 10m ppt-agent-tools
```

---

### Managing lifecycle (start/stop/restart/remove)

- Stop, start, restart:
```bash
docker stop ppt-agent-tools
docker start ppt-agent-tools
docker restart ppt-agent-tools
```

- Remove stopped container and optionally its volumes:
```bash
docker rm ppt-agent-tools
# remove including anonymous volumes
docker rm -v ppt-agent-tools
```

- If an image was retagged/pushed, remove image with:
```bash
docker rmi jeremiahwong/ppt-agent-tools:latest
```

---

### Practical tips & hardening (concise)

- Prefer docker exec for interactive access rather than docker attach.
- Avoid running containers as root in production; add a non-root user in the Dockerfile and use `--user` for exec/run when possible.
- Do not use --privileged unless absolutely necessary. Limit capabilities with `--cap-drop`/`--cap-add`.
- Use bind-mounting for active development so you can edit files locally and test inside the container.
- Persist data with named volumes when state must survive container recreation: `-v ppt-data:/path/in/container`.
- Use a process manager (supervisord or tini) or run the real server CMD in the image instead of `tail -f /dev/null` for production containers.

---

### Quick checklist to get started right now

- [ ] Build and tag image: docker build -t jeremiahwong/ppt-agent-tools:latest -f docker/Dockerfile .
- [ ] Start detached with a name: use the `docker run -d --name ...` example above.
- [ ] Confirm running: docker ps
- [ ] Shell into container: docker exec -it ppt-agent-tools /bin/bash
- [ ] Tail logs for runtime behavior: docker logs -f ppt-agent-tools

---

https://copilot.microsoft.com/shares/1AnRhMudAnSvWSREwWgv9  

