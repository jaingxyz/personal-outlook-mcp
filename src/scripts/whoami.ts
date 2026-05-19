import { getMe } from "../graph.js";

async function main(): Promise<void> {
  const me = await getMe();
  console.log(JSON.stringify(me, null, 2));
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
