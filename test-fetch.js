async function testFetch() {
  const response = await fetch('https://reqres.in/api/users/2');
  const data = await response.json();
  console.log(data);
}

testFetch();