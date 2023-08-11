import { Person } from "@microsoft/mgt-react";

function Persons() {
  const userData = ["bikash.bagaria0@cfexperian.com"];
  return (
    <>
      {userData?.map((userId, index) => {
        return (
          <Person
            personQuery={userId}
            key={index}
            showPresence
            personCardInteraction={1}
            view={5}
          />
        );
      })}
    </>
  );
}

export default Persons;
