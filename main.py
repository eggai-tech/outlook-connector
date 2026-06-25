import asyncio

from eggai import eggai_cleanup

from email_agent import create_email, email_agent
from eggai.transport import eggai_set_default_transport, KafkaTransport


async def main():

    await email_agent.start()
    await create_email()

    try:
        await asyncio.Event().wait()
    except (KeyboardInterrupt, asyncio.CancelledError):
        pass

    await eggai_cleanup()


if __name__ == "__main__":
    eggai_set_default_transport(lambda: KafkaTransport())
    asyncio.run(main())
