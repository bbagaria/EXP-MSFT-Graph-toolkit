import { Person } from "@microsoft/mgt-react";

function Persons() {
  const userData = ["bikash.bagaria@experian.com"];
  return (
    <>
      <div className="person-cls">
        {userData?.map((userId, index) => {
          return (
            <Person
              personQuery={userId}
              key={index}
              showPresence
              personCardInteraction={1}
              view={111}
            />
          );
        })}
      </div>
    </>
  );
}

export default Persons;
