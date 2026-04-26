import PdfSignatureEditor, {
  type SignatureOption,
} from "./components/pdf-signature-editor";

const signatures: SignatureOption[] = [
  {
    attributes: {
      role: "lessor",
      type: "personal",
    },
    description: "Primary personal signature",
    id: "signature-nguyen-van-a",
    name: "Nguyen Van A",
    mimeType: "image/svg+xml",
    kind: "image",
    src: "data:image/svg+xml,%3Csvg%20xmlns='http://www.w3.org/2000/svg'%20viewBox='0%200%20480%20180'%3E%3Crect%20width='480'%20height='180'%20fill='none'/%3E%3Cpath%20d='M42%20116c42-54%2080-72%20112-54%2029%2017%2015%2055-17%2052-33-3-32-48%204-70%2048-29%20103%2048%2064%2086-22%2022-42%2011-33-12%2013%2029%2051%2040%2092%2028%2026-8%2041-28%2041-28s-9%2040%2018%2039c35-1%2056-54%2056-54'%20fill='none'%20stroke='%230f766e'%20stroke-width='12'%20stroke-linecap='round'%20stroke-linejoin='round'/%3E%3Ctext%20x='58'%20y='156'%20font-family='Arial,Helvetica,sans-serif'%20font-size='24'%20font-weight='700'%20fill='%23172033'%3ENguyen%20Van%20A%3C/text%3E%3C/svg%3E",
  },
  {
    attributes: {
      role: "tenant",
      type: "personal",
    },
    description: "Secondary personal signature",
    id: "signature-tran-thi-b",
    kind: "image",
    name: "Tran Thi B",
    mimeType: "image/svg+xml",
    src: "data:image/svg+xml,%3Csvg%20xmlns='http://www.w3.org/2000/svg'%20viewBox='0%200%20480%20180'%3E%3Crect%20width='480'%20height='180'%20fill='none'/%3E%3Cpath%20d='M46%2090c72-56%20121-54%20118-14-3%2037-77%2058-83%2026-6-31%2060-65%20111-27%2046%2034%205%2092-27%2068-22-17%2011-65%2045-51%2025%2010%2016%2054-15%2058-35%204-11-55%2031-50%2036%204%2049%2043%2085%2040%2029-2%2046-22%2062-45'%20fill='none'%20stroke='%231f2937'%20stroke-width='11'%20stroke-linecap='round'%20stroke-linejoin='round'/%3E%3Ctext%20x='58'%20y='156'%20font-family='Arial,Helvetica,sans-serif'%20font-size='24'%20font-weight='700'%20fill='%23172033'%3ETran%20Thi%20B%3C/text%3E%3C/svg%3E",
  },
  {
    attributes: {
      role: "company",
      type: "seal",
    },
    description: "Company approval seal",
    id: "signature-company-seal",
    kind: "image",
    name: "Company Authorized",
    mimeType: "image/svg+xml",
    src: "data:image/svg+xml,%3Csvg%20xmlns='http://www.w3.org/2000/svg'%20viewBox='0%200%20480%20180'%3E%3Crect%20width='480'%20height='180'%20fill='none'/%3E%3Cellipse%20cx='153'%20cy='88'%20rx='102'%20ry='58'%20fill='none'%20stroke='%23b42318'%20stroke-width='9'/%3E%3Cellipse%20cx='153'%20cy='88'%20rx='73'%20ry='36'%20fill='none'%20stroke='%23b42318'%20stroke-width='5'/%3E%3Cpath%20d='M279%20113c24-39%2055-59%2092-47%2024%208%2030%2034%207%2046-21%2011-50-3-43-24%207-22%2050-28%2081%2010'%20fill='none'%20stroke='%23b42318'%20stroke-width='10'%20stroke-linecap='round'%20stroke-linejoin='round'/%3E%3Ctext%20x='88'%20y='95'%20font-family='Arial,Helvetica,sans-serif'%20font-size='22'%20font-weight='800'%20fill='%23b42318'%3EAPPROVED%3C/text%3E%3Ctext%20x='58'%20y='156'%20font-family='Arial,Helvetica,sans-serif'%20font-size='24'%20font-weight='700'%20fill='%23172033'%3ECompany%20Authorized%3C/text%3E%3C/svg%3E",
  },
  {
    attributes: {
      field: "fullName",
      type: "text",
    },
    description: "Signer name text label",
    fontWeight: "bold",
    id: "tag-name-nguyen-van-a",
    kind: "text",
    name: "Name",
    value: "Nguyen Van A",
  },
  {
    attributes: {
      field: "signedDate",
      type: "date",
    },
    description: "Signing date text label",
    fontWeight: "bold",
    id: "tag-date-2026-04-20",
    kind: "text",
    name: "Date",
    value: "20/04/2026",
  },
];

export default function Home() {
  return (
    <PdfSignatureEditor
      options={{
        documentUploadLabel: "Import PDF / DOCX",
        maxPlacements: 300,
        signaturesTitle: "Available signatures",
        title: "PDF Sign Tag",
      }}
      signatures={signatures}
    />
  );
}
