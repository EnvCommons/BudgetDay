import asyncio
import json
import os

from openai import AsyncOpenAI
from openreward import AsyncOpenReward

MODEL_NAME = os.environ.get("MODEL_NAME", "gpt-5.2")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")


async def main() -> None:
    or_client = AsyncOpenReward()
    oai_client = AsyncOpenAI(api_key=OPENAI_API_KEY)

    environment = or_client.environments.get(
        name="GeneralReasoning/BudgetDay"
    )

    tasks = await environment.list_tasks(split="train")
    tools = await environment.list_tools(format="openai")

    print(f"Found {len(tasks)} task(s)")
    print(f"Found {len(tools)} tools total")
    print(f"Tasks: {[t.task_spec for t in tasks]}")
    print()

    if not tasks:
        print("No tasks available!")
        return

    task = tasks[0]
    print(f"Running task: {task.task_spec}")
    print("=" * 80)
    print()

    # Create rollout for tracking this run
    rollout = or_client.rollout.create(
        run_name="budgetday_test",
        rollout_name="test_run",
        environment="local/budgetday",
        split="test",
        task_spec=task.task_spec
    )

    finished = False

    async with environment.session(
        task=task,
        secrets={"openai_api_key": OPENAI_API_KEY}
    ) as session:
        prompt = await session.get_prompt()
        print("PROMPT:")

        # Handle both string and list[TextBlock] prompt formats
        if isinstance(prompt, str):
            prompt_text = prompt
            print(prompt_text)
        else:
            prompt_text = prompt[0].text
            for block in prompt:
                print(block.text)

        print("=" * 80)
        print()

        input_list = [{"role": "user", "content": prompt_text}]

        # Log initial user message to rollout
        rollout.log_openai_response(message=input_list[0], is_finished=False)

        turn_count = 0
        max_turns = 50  # Budget analysis might need more turns

        while not finished and turn_count < max_turns:
            turn_count += 1
            print(f"\n{'='*80}")
            print(f"Turn {turn_count}/{max_turns}")
            print(f"{'='*80}")

            response = await oai_client.responses.create(
                model=MODEL_NAME,
                reasoning={"effort": "medium"},
                tools=tools,
                input=input_list,
            )

            # Add all model outputs to input list
            input_list += response.output

            # Log the model's response
            rollout.log_openai_response(response.output[-1])

            for item in response.output:
                if item.type == "function_call":
                    print(f"\n🔧 Tool Call: {item.name}")
                    print(f"Arguments: {item.arguments}")

                    tool_result = await session.call_tool(
                        item.name,
                        json.loads(str(item.arguments)),
                    )

                    finished = tool_result.finished

                    # Show tool output
                    output_text = tool_result.blocks[0].text if tool_result.blocks else ""
                    print(f"\n📤 Tool Output:")
                    print(output_text[:500] + ("..." if len(output_text) > 500 else ""))  # Truncate long outputs

                    # Feed tool result back to model with correct format
                    result_msg = {
                        "type": "function_call_output",
                        "call_id": item.call_id,
                        "output": output_text,
                    }
                    input_list.append(result_msg)

                    # Log tool result with reward and finished status
                    rollout.log_openai_response(
                        input_list[-1],
                        reward=tool_result.reward,
                        is_finished=finished
                    )

                    reward = tool_result.reward if tool_result.reward is not None else 0.0
                    print(f"\n💰 Reward: {reward:.3f}")
                    if tool_result.finished:
                        print("✅ FINISHED!")
                        finished = True
                        break

                elif item.type == "text":
                    print(f"\n💭 Model Response:")
                    print(item.text[:300] + ("..." if len(item.text) > 300 else ""))  # Truncate long text

        if turn_count >= max_turns:
            print(f"\n⚠️ Reached max turns ({max_turns})")

        print(f"\n{'='*80}")
        print(f"Session Complete - Total Turns: {turn_count}")
        print(f"{'='*80}")


if __name__ == "__main__":
    asyncio.run(main())
