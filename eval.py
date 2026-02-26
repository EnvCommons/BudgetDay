import asyncio
import json
import os
from datetime import datetime
from pathlib import Path

from openai import AsyncOpenAI
from openreward import AsyncOpenReward

MODEL_NAME = os.environ.get("MODEL_NAME", "gpt-5.2")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")


async def run_single_task(
    or_client,
    oai_client,
    environment,
    tools,
    task,
    task_idx: int,
    total_tasks: int,
) -> dict:
    """Run a single task and return results."""
    task_id = task.task_spec.get("task_id", "unknown")

    print(f"\n{'='*80}")
    print(f"Task {task_idx + 1}/{total_tasks}: {task_id}")
    print(f"{'='*80}\n")

    # Create rollout for tracking
    rollout = or_client.rollout.create(
        run_name=f"budgetday_eval",
        rollout_name=task_id,
        environment="local/budgetday",
        split="train",
        task_spec=task.task_spec,
    )

    result = {
        "task_id": task_id,
        "task_type": task.task_spec.get("task_type", "unknown"),
        "status": "not_started",
        "reward": 0.0,
        "turns": 0,
        "error": None,
    }

    try:
        finished = False
        async with environment.session(
            task=task, secrets={"openai_api_key": OPENAI_API_KEY}
        ) as session:
            prompt = await session.get_prompt()

            # Handle both string and list[TextBlock] prompt formats
            if isinstance(prompt, str):
                prompt_text = prompt
            else:
                prompt_text = prompt[0].text

            print(f"📝 Prompt length: {len(prompt_text)} characters\n")

            input_list = [{"role": "user", "content": prompt_text}]
            rollout.log_openai_response(message=input_list[0], is_finished=False)

            turn_count = 0
            max_turns = 200

            while not finished and turn_count < max_turns:
                turn_count += 1
                print(f"Turn {turn_count}/{max_turns}...", end=" ", flush=True)

                response = await oai_client.responses.create(
                    model=MODEL_NAME,
                    reasoning={"effort": "medium"},
                    tools=tools,
                    input=input_list,
                )

                input_list += response.output
                rollout.log_openai_response(response.output[-1])

                # Process tool calls
                for item in response.output:
                    if item.type == "function_call":
                        print(f"🔧 {item.name}", end=" ")

                        tool_result = await session.call_tool(
                            item.name,
                            json.loads(str(item.arguments)),
                        )

                        finished = tool_result.finished
                        output_text = (
                            tool_result.blocks[0].text if tool_result.blocks else ""
                        )

                        result_msg = {
                            "type": "function_call_output",
                            "call_id": item.call_id,
                            "output": output_text,
                        }
                        input_list.append(result_msg)

                        rollout.log_openai_response(
                            input_list[-1],
                            reward=tool_result.reward,
                            is_finished=finished,
                        )

                        if tool_result.reward is not None:
                            result["reward"] = tool_result.reward
                            print(f"💰 {tool_result.reward:.3f}", end=" ")

                        if tool_result.finished:
                            print("✅", end=" ")
                            finished = True
                            break

                print()  # New line after turn

            result["turns"] = turn_count
            result["status"] = "completed" if finished else "max_turns_reached"

            if turn_count >= max_turns:
                print(f"⚠️  Reached max turns ({max_turns})")

    except Exception as e:
        result["status"] = "error"
        result["error"] = str(e)
        print(f"❌ Error: {e}")

    # Print result summary
    status_emoji = {
        "completed": "✅",
        "max_turns_reached": "⏱️",
        "error": "❌",
    }.get(result["status"], "❓")

    print(f"\n{status_emoji} Result: {result['status']}")
    print(f"   Reward: {result['reward']:.3f}")
    print(f"   Turns: {result['turns']}")
    if result["error"]:
        print(f"   Error: {result['error']}")

    return result


async def main() -> None:
    """Run evaluation on all tasks."""
    # Check for API key
    if not OPENAI_API_KEY:
        print("❌ Error: OPENAI_API_KEY environment variable not set")
        return

    # Setup clients
    or_client = AsyncOpenReward()
    oai_client = AsyncOpenAI(api_key=OPENAI_API_KEY)

    # Get environment and tasks
    environment = or_client.environments.get(
        name="GeneralReasoning/BudgetDay"
    )

    tasks = await environment.list_tasks(split="train")
    tools = await environment.list_tools(format="openai")

    print(f"\n{'='*80}")
    print(f"BudgetDay Evaluation")
    print(f"{'='*80}")
    print(f"Model: {MODEL_NAME}")
    print(f"Total tasks: {len(tasks)}")
    print(f"Total tools: {len(tools)}")
    print(f"{'='*80}\n")

    if not tasks:
        print("❌ No tasks available!")
        return

    # Run all tasks
    results = []
    for idx, task in enumerate(tasks):
        result = await run_single_task(
            or_client, oai_client, environment, tools, task, idx, len(tasks)
        )
        results.append(result)

    # Print summary
    print(f"\n{'='*80}")
    print("EVALUATION SUMMARY")
    print(f"{'='*80}\n")

    # Overall stats
    completed = sum(1 for r in results if r["status"] == "completed")
    errors = sum(1 for r in results if r["status"] == "error")
    max_turns = sum(1 for r in results if r["status"] == "max_turns_reached")
    avg_reward = sum(r["reward"] for r in results) / len(results) if results else 0
    avg_turns = sum(r["turns"] for r in results) / len(results) if results else 0

    print(f"Total tasks: {len(results)}")
    print(f"Completed: {completed} ({completed/len(results)*100:.1f}%)")
    print(f"Max turns reached: {max_turns} ({max_turns/len(results)*100:.1f}%)")
    print(f"Errors: {errors} ({errors/len(results)*100:.1f}%)")
    print(f"Average reward: {avg_reward:.3f}")
    print(f"Average turns: {avg_turns:.1f}")
    print()

    # Results by task type
    task_types = {}
    for r in results:
        task_type = r["task_type"]
        if task_type not in task_types:
            task_types[task_type] = []
        task_types[task_type].append(r)

    print("Results by task type:")
    for task_type, type_results in sorted(task_types.items()):
        avg_reward_type = sum(r["reward"] for r in type_results) / len(type_results)
        completed_type = sum(1 for r in type_results if r["status"] == "completed")
        print(f"  {task_type}: {completed_type}/{len(type_results)} completed, avg reward: {avg_reward_type:.3f}")
    print()

    # Detailed results table
    print("Detailed Results:")
    print(f"{'Task ID':<40} {'Type':<15} {'Status':<20} {'Reward':>8} {'Turns':>6}")
    print("-" * 95)
    for r in results:
        status_emoji = {
            "completed": "✅",
            "max_turns_reached": "⏱️",
            "error": "❌",
        }.get(r["status"], "❓")
        print(
            f"{r['task_id']:<40} {r['task_type']:<15} {status_emoji} {r['status']:<17} {r['reward']:>8.3f} {r['turns']:>6}"
        )

    # Save results to JSON
    output_file = f"eval_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    with open(output_file, "w") as f:
        json.dump(
            {
                "model": MODEL_NAME,
                "timestamp": datetime.now().isoformat(),
                "total_tasks": len(results),
                "completed": completed,
                "max_turns_reached": max_turns,
                "errors": errors,
                "average_reward": avg_reward,
                "average_turns": avg_turns,
                "results": results,
            },
            f,
            indent=2,
        )

    print(f"\n📊 Results saved to: {output_file}")
    print(f"{'='*80}\n")


if __name__ == "__main__":
    asyncio.run(main())
