Office.onReady(() => {
  document.getElementById('fetchUserBtn').addEventListener('click', fetchAndInsertData);
});

async function fetchAndInsertData() {
  try {
    const response = await fetch('https://jsonplaceholder.typicode.com/users/1');
    const user = await response.json();

    await Word.run(async (context) => {
      const contentControls = context.document.contentControls;
      contentControls.load("items/tag");
      await context.sync();

      contentControls.items.forEach((cc) => {
        if (cc.tag === "name") {
          cc.insertText(user.name, Word.InsertLocation.replace);
        } else if (cc.tag === "username") {
          cc.insertText(user.username, Word.InsertLocation.replace);
        } else if (cc.tag === "email") {
          cc.insertText(user.email, Word.InsertLocation.replace);
        } else if (cc.tag === "street") {
          cc.insertText(user.address.street, Word.InsertLocation.replace);
        } else if (cc.tag === "city") {
          cc.insertText(user.address.city, Word.InsertLocation.replace);
        } else if (cc.tag === "website") {
          cc.insertText(user.website, Word.InsertLocation.replace);
        }
      });

      await context.sync();
    });
  } catch (error) {
    console.error(error);
    await Word.run(async (context) => {
      context.document.body.insertText('Failed to fetch user data.', Word.InsertLocation.end);
      await context.sync();
    });
  }
}
