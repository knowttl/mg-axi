# Changelog

## [0.1.2](https://github.com/knowttl/mg-axi/compare/v0.1.1...v0.1.2) (2026-10-06)


### Features

* add local-only session-start hook with setup installer ([#63](https://github.com/knowttl/mg-axi/issues/63)) ([6819bc9](https://github.com/knowttl/mg-axi/commit/6819bc926bf014345afb4dcf9dbb36ffb23776b9))
* add read-only Entra delegated-admin count reads ([#66](https://github.com/knowttl/mg-axi/issues/66)) ([99f1efb](https://github.com/knowttl/mg-axi/commit/99f1efb0121a7753f0ae5863d6edaeda3a8d75bf))
* add read-only Entra group PIM schedule reads ([#67](https://github.com/knowttl/mg-axi/issues/67)) ([4e55ba3](https://github.com/knowttl/mg-axi/commit/4e55ba30b744cfaffc384c0b66878c73133cc12e))
* add read-only Entra lifecycle run-nested processing-result reads (#EXT-02i) ([#61](https://github.com/knowttl/mg-axi/issues/61)) ([f70e8a4](https://github.com/knowttl/mg-axi/commit/f70e8a496a936b9b4af2ac4682a5793acb4fccb9))
* add read-only Entra risk-prevention fraud, WAF provider and verification reads ([#59](https://github.com/knowttl/mg-axi/issues/59)) ([3b271b1](https://github.com/knowttl/mg-axi/commit/3b271b1ef81dce6005442f73736454ccc8aa1260))
* add read-only workload risk reads for risky service principals and SP detections ([#65](https://github.com/knowttl/mg-axi/issues/65)) ([85cf5dc](https://github.com/knowttl/mg-axi/commit/85cf5dcdc7515579e5b17c5fe67e98b5dc48624d))
* report uniform N of M totals on core-directory list output ([#68](https://github.com/knowttl/mg-axi/issues/68)) ([2bb7802](https://github.com/knowttl/mg-axi/commit/2bb78027d36579d506f882a66f2d72f14c7fa580))
* report uniform N of M totals on governance list output ([#71](https://github.com/knowttl/mg-axi/issues/71)) ([f3f683e](https://github.com/knowttl/mg-axi/commit/f3f683e99ce2a330ec190b3c8d64f15e5add0c94))
* report uniform N of M totals on group list output ([#64](https://github.com/knowttl/mg-axi/issues/64)) ([483bb4a](https://github.com/knowttl/mg-axi/commit/483bb4a4e2c824062d26d19e3c8fafd8f9bef321))


### Bug Fixes

* align api get truncation with the 500-char rule used elsewhere ([#62](https://github.com/knowttl/mg-axi/issues/62)) ([36a799d](https://github.com/knowttl/mg-axi/commit/36a799d947b6a26fd04a5c91e5b309c92cb888c7))
* cap node test runner at 4 concurrent files ([#70](https://github.com/knowttl/mg-axi/issues/70)) ([306355a](https://github.com/knowttl/mg-axi/commit/306355a4f7fa0543bded0255efc8182518ae15ea))

## [0.1.1](https://github.com/knowttl/mg-axi/compare/v0.1.0...v0.1.1) (2026-10-06)


### Features

* add access-review decision, reviewer, and stage reads ([#32](https://github.com/knowttl/mg-axi/issues/32)) ([0a11b34](https://github.com/knowttl/mg-axi/commit/0a11b34a9d0bf20edeb2e1c28a5f670304c00c73))
* add certificate and workload-federated application authentication ([#5](https://github.com/knowttl/mg-axi/issues/5)) ([8e2dc81](https://github.com/knowttl/mg-axi/commit/8e2dc818d6bc8565cac7523981feb3336a9ec81a))
* add certificate authentication configuration reads ([#35](https://github.com/knowttl/mg-axi/issues/35)) ([4777e0e](https://github.com/knowttl/mg-axi/commit/4777e0e8d9d18a5713610f5fabe9550948fb23e8))
* add Conditional Access policy and named-location reads ([#13](https://github.com/knowttl/mg-axi/issues/13)) ([257134a](https://github.com/knowttl/mg-axi/commit/257134a6269cad6ab60dc7ef4bfb7d877523536c))
* add delegated authentication and profiles ([#4](https://github.com/knowttl/mg-axi/issues/4)) ([b5b6e33](https://github.com/knowttl/mg-axi/commit/b5b6e3336d009d984db09b445c0fdf4fe6eb61f6))
* add directory device and administrative unit reads ([#15](https://github.com/knowttl/mg-axi/issues/15)) ([62e50b9](https://github.com/knowttl/mg-axi/commit/62e50b910c675992c33358fc3cb55be6c4d29c42))
* add Entra application and service-principal reads ([#12](https://github.com/knowttl/mg-axi/issues/12)) ([a0acb36](https://github.com/knowttl/mg-axi/commit/a0acb360680a2e292774e3acdd67850f37b6efc8))
* add Entra authentication-method and registration reads ([#18](https://github.com/knowttl/mg-axi/issues/18)) ([32e8232](https://github.com/knowttl/mg-axi/commit/32e82326ec7b8d46b81172aac6e6d5618b88c394))
* add Entra directory role and PIM reads ([#16](https://github.com/knowttl/mg-axi/issues/16)) ([b19ecaf](https://github.com/knowttl/mg-axi/commit/b19ecaf0c748f350380d097c430874430af35b7e))
* add Entra domain and DNS record reads ([#26](https://github.com/knowttl/mg-axi/issues/26)) ([c718d70](https://github.com/knowttl/mg-axi/commit/c718d70e8199dde6aea072e3f0cced7e83389a45))
* add Entra group and membership reads ([#10](https://github.com/knowttl/mg-axi/issues/10)) ([78bcd84](https://github.com/knowttl/mg-axi/commit/78bcd846eac8ea0a775e3b38c396b0890b0fe40b))
* add Entra group lifecycle policy and setting template reads ([#34](https://github.com/knowttl/mg-axi/issues/34)) ([3092035](https://github.com/knowttl/mg-axi/commit/3092035e6b37c33979903cc384c07209b109bc92))
* add Entra organization and branding reads ([#27](https://github.com/knowttl/mg-axi/issues/27)) ([0af10a0](https://github.com/knowttl/mg-axi/commit/0af10a0a3a1998207ee0ae2a1c8bdb80189e893e))
* add Entra risky-user and risk-detection reads ([#14](https://github.com/knowttl/mg-axi/issues/14)) ([e94f2f8](https://github.com/knowttl/mg-axi/commit/e94f2f8dec673301d538e6c44064baba70f674c8))
* add Entra sign-in and directory audit log reads ([#11](https://github.com/knowttl/mg-axi/issues/11)) ([6bc0c3f](https://github.com/knowttl/mg-axi/commit/6bc0c3fff4b04671b41fb3e0bcca0fca415cbcd7))
* add Entra user list and show commands ([#8](https://github.com/knowttl/mg-axi/issues/8)) ([c51cab9](https://github.com/knowttl/mg-axi/commit/c51cab96ec66f9a46de4e4d7f080288be7aeda6c))
* add fixture-only mutation coordinator and execution gates ([#20](https://github.com/knowttl/mg-axi/issues/20)) ([01d3b45](https://github.com/knowttl/mg-axi/commit/01d3b4594a4e3615f2a43b67513d627e0f4d130d))
* add gated Conditional Access policy updates ([#25](https://github.com/knowttl/mg-axi/issues/25)) ([8961172](https://github.com/knowttl/mg-axi/commit/8961172a24d4e14e85467c9aa8caf0a48c88ac18))
* add gated Entra account enable and disable ([#22](https://github.com/knowttl/mg-axi/issues/22)) ([b0291ec](https://github.com/knowttl/mg-axi/commit/b0291ecee207e01e17a690126eaaf96f73790b34))
* add gated Entra user group membership writes ([#21](https://github.com/knowttl/mg-axi/issues/21)) ([d46c6bf](https://github.com/knowttl/mg-axi/commit/d46c6bf2105b5974caa9952052b7d1430187aea5))
* add gated Entra user session revocation ([#23](https://github.com/knowttl/mg-axi/issues/23)) ([dd7c046](https://github.com/knowttl/mg-axi/commit/dd7c04661675a2f1b34a0c2af61af76c3a5f158b))
* add gated single-user risk dismissal ([#24](https://github.com/knowttl/mg-axi/issues/24)) ([74dbc88](https://github.com/knowttl/mg-axi/commit/74dbc88e328cea1aac8a1a044b28a80781af89c4))
* add pinned Entra operation inventory and offline tooling ([#2](https://github.com/knowttl/mg-axi/issues/2)) ([bd93770](https://github.com/knowttl/mg-axi/commit/bd93770a81f544f7a5547a0e3c00b93dc8939365))
* add policy-enforced Graph read session ([#6](https://github.com/knowttl/mg-axi/issues/6)) ([50bcb99](https://github.com/knowttl/mg-axi/commit/50bcb9944ec036fbfc1f6143563a9be702a86c18))
* add read-only attribute-set and custom-security-attribute reads ([#36](https://github.com/knowttl/mg-axi/issues/36)) ([feee1ba](https://github.com/knowttl/mg-axi/commit/feee1bac5dfeeb34598289c77b1eee0b09d37922))
* add read-only Entra access review commands ([#28](https://github.com/knowttl/mg-axi/issues/28)) ([029a69a](https://github.com/knowttl/mg-axi/commit/029a69a49483c6a2e0fe393aef3e12d16b9dfa6c))
* add read-only Entra commercial subscription commands ([#37](https://github.com/knowttl/mg-axi/issues/37)) ([4da700f](https://github.com/knowttl/mg-axi/commit/4da700f3dfaf7113bb3a89571fefba77ae2013ef))
* add read-only Entra contact manager and direct-report navigation reads ([#46](https://github.com/knowttl/mg-axi/issues/46)) ([f963008](https://github.com/knowttl/mg-axi/commit/f96300880d9dfeeb9861694c12f050c816b3567f))
* add read-only Entra contact memberOf/transitiveMemberOf membership reads ([#47](https://github.com/knowttl/mg-axi/issues/47)) ([e18896c](https://github.com/knowttl/mg-axi/commit/e18896cd6a33f7413def85c06caf3f6631ac571d))
* add read-only Entra data-policy-operation commands ([#41](https://github.com/knowttl/mg-axi/issues/41)) ([06155b7](https://github.com/knowttl/mg-axi/commit/06155b76b73681b253863829964edada37bd5ce8))
* add read-only Entra delegated-admin access-assignment, operation, request and service-management-detail reads ([#50](https://github.com/knowttl/mg-axi/issues/50)) ([d9bc546](https://github.com/knowttl/mg-axi/commit/d9bc546737ad5eb6dfb4d716bf573f32c6aaecc9))
* add read-only Entra delegated-admin customer and relationship reads ([#48](https://github.com/knowttl/mg-axi/issues/48)) ([2d6ffb1](https://github.com/knowttl/mg-axi/commit/2d6ffb16cb0b86d9750f9e97d0563072a0b86612))
* add read-only Entra deleted directory-item commands (#EXT-01l) ([#45](https://github.com/knowttl/mg-axi/issues/45)) ([76bc025](https://github.com/knowttl/mg-axi/commit/76bc025d8a315d3a62ecd51936e3a415b12d71ca))
* add read-only Entra directory federation-configuration commands ([#43](https://github.com/knowttl/mg-axi/issues/43)) ([9baafca](https://github.com/knowttl/mg-axi/commit/9baafca8957949db8a3a84ab1ff697bfd2b54e64))
* add read-only Entra directory-object list/show/count commands ([#42](https://github.com/knowttl/mg-axi/issues/42)) ([c6655d2](https://github.com/knowttl/mg-axi/commit/c6655d27cdc15aee8c6edce1756280e996d7171a))
* add read-only Entra entitlement assignment and assignment-request reads ([#53](https://github.com/knowttl/mg-axi/issues/53)) ([72e4057](https://github.com/knowttl/mg-axi/commit/72e40576de4b25d4e8f175d767e8979d83ae2a10))
* add read-only Entra entitlement catalog, access-package, assignment-policy and resource-role-scope reads ([#52](https://github.com/knowttl/mg-axi/issues/52)) ([7811c48](https://github.com/knowttl/mg-axi/commit/7811c48fcd190dc82be4b5168eca28c9ce489c17))
* add read-only Entra lifecycle run and processing-result reads ([#55](https://github.com/knowttl/mg-axi/issues/55)) ([50f22dc](https://github.com/knowttl/mg-axi/commit/50f22dc12e9840f41336e8caea75b954c9e15e1c))
* add read-only Entra lifecycle task-report reads (#EXT-02h) ([#56](https://github.com/knowttl/mg-axi/issues/56)) ([5f1b482](https://github.com/knowttl/mg-axi/commit/5f1b4823dc22b8fb953b2d45ddba6d9e6dc47741))
* add read-only Entra lifecycle workflow, template, task-definition and settings reads (#EXT-02f) ([#54](https://github.com/knowttl/mg-axi/issues/54)) ([ba31377](https://github.com/knowttl/mg-axi/commit/ba31377b447bef744ede47ebad92f646f7b78e80))
* add read-only Entra multi-tenant-organization show/list/count reads ([#49](https://github.com/knowttl/mg-axi/issues/49)) ([b6b4530](https://github.com/knowttl/mg-axi/commit/b6b4530b2788862096f1fb11e74a200068cbc79d))
* add read-only Entra on-premises-synchronization commands ([#39](https://github.com/knowttl/mg-axi/issues/39)) ([a260d85](https://github.com/knowttl/mg-axi/commit/a260d8587030942050eb9fb7a406619781dcc18a))
* add read-only Entra organizational-contact list/show/count commands ([#44](https://github.com/knowttl/mg-axi/issues/44)) ([1258657](https://github.com/knowttl/mg-axi/commit/1258657caa76e292402eabb1ac4c758e8da78ac1))
* add read-only Entra terms-of-use agreement and acceptance commands ([#40](https://github.com/knowttl/mg-axi/issues/40)) ([b8eb224](https://github.com/knowttl/mg-axi/commit/b8eb224e98ea9636445537d9a8094b525bd8b2c2))
* add read-only partner contract commands ([#30](https://github.com/knowttl/mg-axi/issues/30)) ([44152e6](https://github.com/knowttl/mg-axi/commit/44152e66453fc2f1cea52a12ec33c0646f8192fa))
* add resumable Graph collections and bounded read retries ([#7](https://github.com/knowttl/mg-axi/issues/7)) ([92bcb36](https://github.com/knowttl/mg-axi/commit/92bcb36b965d2d47ba5a49337f532ae35059cdcf))
* add reviewed read-only Graph API command ([#9](https://github.com/knowttl/mg-axi/issues/9)) ([bebcaa7](https://github.com/knowttl/mg-axi/commit/bebcaa75b79a63c6e20e6c277b1c594313f7d6f5))
* add service-principal consent grant reads ([#17](https://github.com/knowttl/mg-axi/issues/17)) ([a0c8588](https://github.com/knowttl/mg-axi/commit/a0c8588f8d089ce7c3b5f007f6a63c0f1d60b86b))
* add setup, doctor and packaged Entra read documentation ([#19](https://github.com/knowttl/mg-axi/issues/19)) ([185dde9](https://github.com/knowttl/mg-axi/commit/185dde96c4f9b8a4bafe9062efea12a457b82f2c))
* add tenant-information and commerce-key subscription reads via validated function binding ([#51](https://github.com/knowttl/mg-axi/issues/51)) ([b581e6f](https://github.com/knowttl/mg-axi/commit/b581e6fdd384f252fbc57b7a477ac072e3c2e3c1))
* add the initial TypeScript AXI CLI shell ([#3](https://github.com/knowttl/mg-axi/issues/3)) ([cb60929](https://github.com/knowttl/mg-axi/commit/cb609297c65518a6964843e865fc97d2200abd28))
* add workforce identity-provider reads ([#29](https://github.com/knowttl/mg-axi/issues/29)) ([26dd7e0](https://github.com/knowttl/mg-axi/commit/26dd7e034aee2435ee3a3882bdd914b311c1c4d5))
* publish @knowttl/mg-axi to npm with release-please and npx-based install ([#57](https://github.com/knowttl/mg-axi/issues/57)) ([51004cc](https://github.com/knowttl/mg-axi/commit/51004cc9de4150c9929fa07c0e8ffb82bb057a39))


### Bug Fixes

* add mg-axi entra contact list/show/count on the shared session with OrgContact.Read.All default (added to the shared READ_SCOPES allowlist; it is a read scope so nothing is deferred), minimal personal-data defaults (id, displayName, mail, companyName) with flat-scalar-only projection and no $expand, --filter with the documented $count=true plus ConsistencyLevel eventual contract, ConsistencyLevel eventual on the $count scalar, v1.0-only gating with beta refused before credentials, and two reviewed raw routes with the same field set. Defer the v1.0 delta() sync, the per-contact navigation reads, the POST lookup/validation actions and the beta operations via a tools/inventory.py rule plus regeneration. ([1258657](https://github.com/knowttl/mg-axi/commit/1258657caa76e292402eabb1ac4c758e8da78ac1))
* add mg-axi entra directory-object list/show/count on the shared session with Directory.Read.All default (already in the shared READ_SCOPES allowlist; it is a read scope so nothing is deferred), base-properties-only projection with the [@odata](https://github.com/odata).type discriminator riding along automatically, $select-only lists with no --filter, v1.0-only gating with beta refused before credentials, and two reviewed raw routes that likewise keep the discriminator. Defer the v1.0 delta() sync, the POST lookup/validation actions and the beta operations via a tools/inventory.py rule plus regeneration. ([c6655d2](https://github.com/knowttl/mg-axi/commit/c6655d27cdc15aee8c6edce1756280e996d7171a))
* **inventory:** mark undocumented v1.0 invitation reads unavailable ([#31](https://github.com/knowttl/mg-axi/issues/31)) ([a32f663](https://github.com/knowttl/mg-axi/commit/a32f6630c0992a553e5b49b943da9bf58dcae7d7))
* keep mg-axi agent guidance aligned with the command catalogue ([#33](https://github.com/knowttl/mg-axi/issues/33)) ([0d578c9](https://github.com/knowttl/mg-axi/commit/0d578c99be2128ef8b499ce743f4adb2a7070a10))
